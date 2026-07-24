using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace EmailNotes
{
    /// <summary>
    /// The docked notepad shown next to the mail. A plain WinForms control so it
    /// loads locally with no login, no web view and no internet. It polls the
    /// Outlook selection (late-bound) and keeps a note per mail.
    /// </summary>
    [ComVisible(true)]
    [Guid("B1D9E7C2-6F1A-4C2E-9E7D-2A5B3C4D5E61")]
    [ProgId("EmailNotes.NotesPane")]
    [ClassInterface(ClassInterfaceType.None)]
    public class NotesPane : UserControl
    {
        private readonly Store _store = new Store();

        private Panel _updateBar;
        private Button _updateLink;
        private Updater.Info _update;

        private Label _header;
        private TextBox _note;
        private Button _btnShot;
        private Button _btnFile;
        private ListBox _attList;
        private Button _btnOpen;
        private Button _btnRemove;
        private Label _saved;

        private readonly Timer _poll = new Timer();
        private readonly Timer _saveTimer = new Timer();

        private object _app;
        private string _mailKey;
        private Note _current;

        public NotesPane()
        {
            BuildUi();
            _poll.Interval = 800;
            _poll.Tick += (s, e) => Poll();
            _saveTimer.Interval = 600;
            _saveTimer.Tick += (s, e) => { _saveTimer.Stop(); FlushSave(); };
        }

        /// <summary>Called by the add-in once the Outlook Application is known.</summary>
        public void Initialize(object outlookApplication)
        {
            _app = outlookApplication;
            ShowNoSelection();
            _poll.Start();
            Poll();
            // Offer an update at startup if a newer release exists (best effort, offline-safe).
            System.Threading.ThreadPool.QueueUserWorkItem(_ =>
            {
                var info = Updater.Check();
                if (info == null) return;
                try { if (IsHandleCreated) BeginInvoke((Action)(() => ShowUpdateBar(info))); } catch { }
            });
        }

        private void ShowUpdateBar(Updater.Info info)
        {
            _update = info;
            _updateLink.Text = "⬆  Neue Version " + info.Version + " verfügbar – jetzt aktualisieren";
            _updateBar.Visible = true;
        }

        private void OnUpdateClick()
        {
            if (_update == null) return;
            _updateLink.Text = "Lädt Update…";
            System.Threading.ThreadPool.QueueUserWorkItem(_ => Updater.DownloadAndRun(_update));
        }

        // ---- UI ----

        private void BuildUi()
        {
            BackColor = Color.FromArgb(31, 29, 36);
            ForeColor = Color.FromArgb(236, 236, 240);
            Font = new Font("Segoe UI", 9f);
            Padding = new Padding(10);

            _header = new Label
            {
                Dock = DockStyle.Top,
                AutoSize = false,
                Height = 46,
                Text = "E-Mail-Notizen",
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };

            var actions = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 36, WrapContents = false, AutoSize = false };
            _btnShot = MakeButton("📋 Screenshot einfügen");
            _btnShot.Click += (s, e) => PasteScreenshot();
            _btnFile = MakeButton("📎 Datei anhängen");
            _btnFile.Click += (s, e) => AttachFiles();
            actions.Controls.Add(_btnShot);
            actions.Controls.Add(_btnFile);

            _note = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                AcceptsReturn = true,
                AcceptsTab = true,
                BackColor = Color.FromArgb(42, 39, 49),
                ForeColor = Color.FromArgb(236, 236, 240),
                BorderStyle = BorderStyle.FixedSingle,
                Enabled = false
            };
            _note.TextChanged += (s, e) => { if (_current != null) { _saveTimer.Stop(); _saveTimer.Start(); SetSaving(); } };
            _note.KeyDown += Note_KeyDown;

            var attGroup = new Panel { Dock = DockStyle.Bottom, Height = 150 };
            var attLabel = new Label { Dock = DockStyle.Top, Height = 20, Text = "Anhänge", ForeColor = Color.FromArgb(167, 163, 179) };
            _attList = new ListBox
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(42, 39, 49),
                ForeColor = Color.FromArgb(236, 236, 240),
                BorderStyle = BorderStyle.FixedSingle
            };
            _attList.DoubleClick += (s, e) => OpenAttachment();
            var attButtons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 34, WrapContents = false };
            _btnOpen = MakeButton("Öffnen");
            _btnOpen.Click += (s, e) => OpenAttachment();
            _btnRemove = MakeButton("Entfernen");
            _btnRemove.Click += (s, e) => RemoveAttachment();
            attButtons.Controls.Add(_btnOpen);
            attButtons.Controls.Add(_btnRemove);
            attGroup.Controls.Add(_attList);
            attGroup.Controls.Add(attButtons);
            attGroup.Controls.Add(attLabel);

            _saved = new Label { Dock = DockStyle.Bottom, Height = 18, ForeColor = Color.FromArgb(167, 163, 179), Text = "" };

            _updateBar = new Panel { Dock = DockStyle.Top, Height = 30, BackColor = Color.FromArgb(124, 92, 246), Visible = false };
            _updateLink = new Button
            {
                Dock = DockStyle.Fill,
                FlatStyle = FlatStyle.Flat,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.White,
                BackColor = Color.FromArgb(124, 92, 246),
                Cursor = Cursors.Hand,
                Text = "Update verfügbar"
            };
            _updateLink.FlatAppearance.BorderSize = 0;
            _updateLink.Click += (s, e) => OnUpdateClick();
            var dismiss = new Button
            {
                Dock = DockStyle.Right,
                Width = 28,
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.White,
                BackColor = Color.FromArgb(124, 92, 246),
                Text = "✕"
            };
            dismiss.FlatAppearance.BorderSize = 0;
            dismiss.Click += (s, e) => { _updateBar.Visible = false; };
            _updateBar.Controls.Add(_updateLink);
            _updateBar.Controls.Add(dismiss);

            // Add in reverse dock order (Fill first, outermost edges last).
            Controls.Add(_note);
            Controls.Add(attGroup);
            Controls.Add(_saved);
            Controls.Add(actions);
            Controls.Add(_header);
            Controls.Add(_updateBar);
        }

        private Button MakeButton(string text)
        {
            return new Button
            {
                Text = text,
                AutoSize = true,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 2, 6, 2),
                BackColor = Color.FromArgb(50, 46, 59),
                ForeColor = Color.FromArgb(236, 236, 240)
            };
        }

        // ---- selection polling ----

        private void Poll()
        {
            try
            {
                dynamic app = _app;
                if (app == null) return;
                dynamic exp = app.ActiveExplorer();
                if (exp == null) { ShowNoSelection(); return; }
                dynamic sel = exp.Selection;
                if (sel == null || (int)sel.Count < 1) { ShowNoSelection(); return; }

                dynamic item = null;
                int count = (int)sel.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic it = sel.Item(i);
                    try { if ((int)it.Class == 43) { item = it; break; } } catch { }
                }
                if (item == null) { ShowNoSelection(); return; }

                string entryId = (string)item.EntryID;
                string subject = Safe(() => (string)item.Subject);
                string sender = Safe(() => (string)item.SenderName);
                OnMail(entryId, subject, sender);
            }
            catch { /* retry next tick */ }
        }

        private static string Safe(Func<string> f) { try { return f() ?? ""; } catch { return ""; } }

        private void OnMail(string entryId, string subject, string sender)
        {
            if (string.IsNullOrEmpty(entryId)) { ShowNoSelection(); return; }
            if (entryId == _mailKey)
            {
                if (_current != null) _header.Text = HeaderText(_current, subject, sender);
                return;
            }
            // Switched mail → persist the previous note first.
            FlushSave();
            _mailKey = entryId;
            _current = _store.GetOrCreate(entryId, subject, sender);
            LoadCurrent(subject, sender);
        }

        private void LoadCurrent(string subject, string sender)
        {
            _header.Text = HeaderText(_current, subject, sender);
            _note.Enabled = true;
            _note.Text = _current.Text ?? "";
            _saved.Text = "wird automatisch gespeichert";
            RefreshAttachments();
        }

        private string HeaderText(Note n, string subject, string sender)
        {
            string subj = string.IsNullOrEmpty(subject) ? (n.Subject ?? "") : subject;
            string snd = string.IsNullOrEmpty(sender) ? (n.Sender ?? "") : sender;
            string line = "Notiz #" + n.NoteId + "  " + (string.IsNullOrEmpty(subj) ? "(ohne Betreff)" : subj);
            if (!string.IsNullOrEmpty(snd)) line += "\nvon " + snd;
            return line;
        }

        private void ShowNoSelection()
        {
            if (_current == null && !_note.Enabled) return;
            FlushSave();
            _mailKey = null;
            _current = null;
            _header.Text = "Keine Mail markiert";
            _note.Enabled = false;
            _note.Text = "";
            _saved.Text = "";
            _attList.Items.Clear();
        }

        // ---- saving ----

        private void SetSaving() { _saved.Text = "Speichere…"; }

        private void FlushSave()
        {
            if (_current == null || _mailKey == null) return;
            _store.SaveText(_mailKey, _note.Text);
            _saved.Text = "gespeichert " + DateTime.Now.ToString("HH:mm:ss");
        }

        // ---- attachments ----

        private void RefreshAttachments()
        {
            _attList.Items.Clear();
            if (_current == null) return;
            foreach (var a in _current.Attachments)
                _attList.Items.Add((a.IsImage ? "🖼 " : "📎 ") + a.Name);
        }

        private string SelectedAttachmentName()
        {
            if (_current == null || _attList.SelectedIndex < 0) return null;
            if (_attList.SelectedIndex >= _current.Attachments.Count) return null;
            return _current.Attachments[_attList.SelectedIndex].Name;
        }

        private void Note_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V && Clipboard.ContainsImage())
            {
                e.SuppressKeyPress = true;
                PasteScreenshot();
            }
        }

        private void PasteScreenshot()
        {
            if (_current == null) return;
            if (!Clipboard.ContainsImage()) { _saved.Text = "Kein Bild in der Zwischenablage."; return; }
            try
            {
                using (Image img = Clipboard.GetImage())
                {
                    string name = _store.NextImageName(_current.NoteId);
                    string path = Path.Combine(_store.AttachmentsDir(_current.NoteId), name);
                    img.Save(path, ImageFormat.Png);
                    _current = _store.AddAttachment(_mailKey, name, true);
                    InsertToken(name);
                    RefreshAttachments();
                }
            }
            catch (Exception ex) { _saved.Text = "Fehler: " + ex.Message; }
        }

        private void AttachFiles()
        {
            if (_current == null) return;
            using (var dlg = new OpenFileDialog { Multiselect = true, Title = "Anhänge hinzufügen" })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;
                foreach (var src in dlg.FileNames)
                {
                    try
                    {
                        string name = _store.UniqueName(_current.NoteId, Path.GetFileName(src));
                        string dst = Path.Combine(_store.AttachmentsDir(_current.NoteId), name);
                        File.Copy(src, dst, false);
                        bool isImg = IsImage(name);
                        _current = _store.AddAttachment(_mailKey, name, isImg);
                        InsertToken(name);
                    }
                    catch { }
                }
                RefreshAttachments();
            }
        }

        private void OpenAttachment()
        {
            string name = SelectedAttachmentName();
            if (name == null) return;
            try { System.Diagnostics.Process.Start(Path.Combine(_store.AttachmentsDir(_current.NoteId), name)); }
            catch { }
        }

        private void RemoveAttachment()
        {
            string name = SelectedAttachmentName();
            if (name == null) return;
            try { File.Delete(Path.Combine(_store.AttachmentsDir(_current.NoteId), name)); } catch { }
            _current = _store.RemoveAttachment(_mailKey, name);
            string token = "[📎 " + name + "]";
            if (_note.Text.Contains(token)) _note.Text = _note.Text.Replace(token, "");
            RefreshAttachments();
        }

        private void InsertToken(string name)
        {
            string token = "[📎 " + name + "]";
            int at = _note.SelectionStart;
            string before = _note.Text.Substring(0, at);
            string pad = (before.Length > 0 && !char.IsWhiteSpace(before[before.Length - 1])) ? " " : "";
            string insert = pad + token + " ";
            _note.Text = before + insert + _note.Text.Substring(at);
            _note.SelectionStart = at + insert.Length;
            _note.Focus();
        }

        private static bool IsImage(string name)
        {
            string ext = Path.GetExtension(name).ToLowerInvariant();
            return ext == ".png" || ext == ".jpg" || ext == ".jpeg" || ext == ".gif" ||
                   ext == ".bmp" || ext == ".webp" || ext == ".tif" || ext == ".tiff";
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                try { FlushSave(); } catch { }
                _poll.Stop();
                _saveTimer.Stop();
            }
            base.Dispose(disposing);
        }
    }
}
