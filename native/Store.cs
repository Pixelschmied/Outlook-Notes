using System;
using System.Collections.Generic;
using System.IO;
using System.Web.Script.Serialization;

namespace EmailNotes
{
    /// <summary>One attachment of a note (the file lives next to the store).</summary>
    public class Attachment
    {
        public string Name { get; set; }
        public bool IsImage { get; set; }
    }

    /// <summary>A note bound to a mail, with a friendly sequential number.</summary>
    public class Note
    {
        public int NoteId { get; set; }
        public string MailKey { get; set; }
        public string Subject { get; set; }
        public string Sender { get; set; }
        public string Text { get; set; }
        public List<Attachment> Attachments { get; set; } = new List<Attachment>();
        public string UpdatedAt { get; set; }
    }

    /// <summary>
    /// Local, offline note storage — everything lives under %APPDATA%\EmailNotes.
    /// Notes are keyed by the mail's Outlook EntryID and given a sequential
    /// number the first time a mail is seen (Mail 1 = Note 1, Mail 2 = Note 2 …).
    /// </summary>
    internal class Store
    {
        private class Model
        {
            public int Counter { get; set; }
            public Dictionary<string, Note> Notes { get; set; } = new Dictionary<string, Note>();
        }

        private static readonly object Gate = new object();
        private readonly string _dir;
        private readonly string _file;
        private readonly JavaScriptSerializer _json = new JavaScriptSerializer();
        private Model _model;

        public Store()
        {
            _dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "EmailNotes");
            Directory.CreateDirectory(_dir);
            _file = Path.Combine(_dir, "notes.json");
            Load();
        }

        public string AttachmentsDir(int noteId)
        {
            string p = Path.Combine(_dir, "attachments", noteId.ToString());
            Directory.CreateDirectory(p);
            return p;
        }

        private void Load()
        {
            try
            {
                if (File.Exists(_file))
                {
                    _model = _json.Deserialize<Model>(File.ReadAllText(_file)) ?? new Model();
                    if (_model.Notes == null) _model.Notes = new Dictionary<string, Note>();
                }
                else _model = new Model();
            }
            catch { _model = new Model(); }
        }

        private void Save()
        {
            try { File.WriteAllText(_file, _json.Serialize(_model)); }
            catch { /* best effort */ }
        }

        /// <summary>Resolve a mail to its note, creating an empty one (next number) on first sight.</summary>
        public Note GetOrCreate(string mailKey, string subject, string sender)
        {
            lock (Gate)
            {
                Note n;
                if (!_model.Notes.TryGetValue(mailKey, out n))
                {
                    _model.Counter++;
                    n = new Note
                    {
                        NoteId = _model.Counter,
                        MailKey = mailKey,
                        Subject = subject ?? "",
                        Sender = sender ?? "",
                        Text = "",
                        UpdatedAt = DateTime.Now.ToString("o")
                    };
                    _model.Notes[mailKey] = n;
                    Save();
                }
                else
                {
                    if (!string.IsNullOrEmpty(subject)) n.Subject = subject;
                    if (!string.IsNullOrEmpty(sender)) n.Sender = sender;
                }
                return n;
            }
        }

        public void SaveText(string mailKey, string text)
        {
            lock (Gate)
            {
                Note n;
                if (_model.Notes.TryGetValue(mailKey, out n))
                {
                    n.Text = text ?? "";
                    n.UpdatedAt = DateTime.Now.ToString("o");
                    Save();
                }
            }
        }

        public Note AddAttachment(string mailKey, string name, bool isImage)
        {
            lock (Gate)
            {
                Note n;
                if (!_model.Notes.TryGetValue(mailKey, out n)) return null;
                n.Attachments.Add(new Attachment { Name = name, IsImage = isImage });
                n.UpdatedAt = DateTime.Now.ToString("o");
                Save();
                return n;
            }
        }

        public Note RemoveAttachment(string mailKey, string name)
        {
            lock (Gate)
            {
                Note n;
                if (!_model.Notes.TryGetValue(mailKey, out n)) return null;
                n.Attachments.RemoveAll(a => a.Name == name);
                n.UpdatedAt = DateTime.Now.ToString("o");
                Save();
                return n;
            }
        }

        /// <summary>Free, collision-safe file name inside a note's attachment folder.</summary>
        public string UniqueName(int noteId, string desired)
        {
            string folder = AttachmentsDir(noteId);
            string stem = Path.GetFileNameWithoutExtension(desired);
            string ext = Path.GetExtension(desired);
            if (string.IsNullOrEmpty(stem)) stem = "Datei";
            string name = stem + ext;
            int i = 1;
            while (File.Exists(Path.Combine(folder, name))) { name = stem + " (" + i + ")" + ext; i++; }
            return name;
        }

        /// <summary>Next sequential screenshot name ("1.png", "2.png", …) for a note.</summary>
        public string NextImageName(int noteId)
        {
            string folder = AttachmentsDir(noteId);
            int n = 1;
            while (File.Exists(Path.Combine(folder, n + ".png"))) n++;
            return n + ".png";
        }
    }
}
