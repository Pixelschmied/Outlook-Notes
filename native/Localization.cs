using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;

namespace OutlookNotes
{
    /// <summary>All user-visible strings for one language.</summary>
    internal sealed class Loc
    {
        public string PaneTitle;
        public string InsertScreenshot;
        public string AttachFile;
        public string Attachments;
        public string Open;
        public string Remove;
        public string AutoSaved;
        public string Saving;
        public string SavedAt;             // {0} = time
        public string NoMailSelected;
        public string NoImageInClipboard;
        public string ErrorPrefix;         // {0} = message
        public string NoteHeader;          // {0} = id, {1} = subject
        public string NoSubject;
        public string From;                // {0} = sender
        public string UpdateAvailable;     // {0} = version
        public string Loading;
        public string UpdateBadge;
        public string AddAttachmentsDialog;
    }

    /// <summary>
    /// Tiny built-in localization. The UI language is taken from Outlook's own
    /// UI language when available (so the notepad matches the rest of Outlook),
    /// otherwise from the Windows display language. Unknown languages fall back
    /// to English. No emoji is used anywhere so nothing renders as a "tofu" box.
    /// </summary>
    internal static class Localization
    {
        /// <summary>Two-letter language of the Windows UI (used before Outlook is known).</summary>
        public static string DetectFromCulture()
        {
            try { return CultureInfo.CurrentUICulture.TwoLetterISOLanguageName.ToLowerInvariant(); }
            catch { return "en"; }
        }

        /// <summary>Two-letter language of Outlook's UI, falling back to the Windows UI language.</summary>
        public static string DetectLanguage(object outlookApp)
        {
            try
            {
                dynamic app = outlookApp;
                // MsoAppLanguageID.msoLanguageIDUI = 2
                object idObj = app.LanguageSettings.LanguageID(2);
                int lcid = Convert.ToInt32(idObj);
                if (lcid > 0)
                    return new CultureInfo(lcid).TwoLetterISOLanguageName.ToLowerInvariant();
            }
            catch { }
            return DetectFromCulture();
        }

        /// <summary>A UI font that can actually render the given language (CJK needs a CJK font).</summary>
        public static Font UiFont(string lang, float size, FontStyle style = FontStyle.Regular)
        {
            string family;
            switch (lang)
            {
                case "ja": family = "Yu Gothic UI"; break;      // Japanese
                case "zh": family = "Microsoft YaHei UI"; break; // Chinese
                case "ko": family = "Malgun Gothic"; break;      // Korean
                default: family = "Segoe UI"; break;
            }
            try { return new Font(family, size, style); }
            catch { return new Font("Segoe UI", size, style); }
        }

        public static Loc Get(string lang)
        {
            Loc l;
            return Table.TryGetValue(lang ?? "en", out l) ? l : Table["en"];
        }

        private static readonly Dictionary<string, Loc> Table = new Dictionary<string, Loc>
        {
            ["en"] = new Loc
            {
                PaneTitle = "Email notes",
                InsertScreenshot = "Insert screenshot",
                AttachFile = "Attach file",
                Attachments = "Attachments",
                Open = "Open",
                Remove = "Remove",
                AutoSaved = "saved automatically",
                Saving = "Saving…",
                SavedAt = "saved {0}",
                NoMailSelected = "No mail selected",
                NoImageInClipboard = "No image in the clipboard.",
                ErrorPrefix = "Error: {0}",
                NoteHeader = "Note #{0}  {1}",
                NoSubject = "(no subject)",
                From = "from {0}",
                UpdateAvailable = "Update {0} available",
                Loading = "Downloading update…",
                UpdateBadge = "Update available",
                AddAttachmentsDialog = "Add attachments",
            },
            ["de"] = new Loc
            {
                PaneTitle = "E-Mail-Notizen",
                InsertScreenshot = "Screenshot einfügen",
                AttachFile = "Datei anhängen",
                Attachments = "Anhänge",
                Open = "Öffnen",
                Remove = "Entfernen",
                AutoSaved = "wird automatisch gespeichert",
                Saving = "Speichere…",
                SavedAt = "gespeichert {0}",
                NoMailSelected = "Keine Mail markiert",
                NoImageInClipboard = "Kein Bild in der Zwischenablage.",
                ErrorPrefix = "Fehler: {0}",
                NoteHeader = "Notiz #{0}  {1}",
                NoSubject = "(ohne Betreff)",
                From = "von {0}",
                UpdateAvailable = "Update {0} verfügbar",
                Loading = "Update wird geladen…",
                UpdateBadge = "Update verfügbar",
                AddAttachmentsDialog = "Anhänge hinzufügen",
            },
            ["fr"] = new Loc
            {
                PaneTitle = "Notes d'e-mail",
                InsertScreenshot = "Insérer une capture",
                AttachFile = "Joindre un fichier",
                Attachments = "Pièces jointes",
                Open = "Ouvrir",
                Remove = "Supprimer",
                AutoSaved = "enregistré automatiquement",
                Saving = "Enregistrement…",
                SavedAt = "enregistré {0}",
                NoMailSelected = "Aucun e-mail sélectionné",
                NoImageInClipboard = "Aucune image dans le presse-papiers.",
                ErrorPrefix = "Erreur : {0}",
                NoteHeader = "Note n°{0}  {1}",
                NoSubject = "(sans objet)",
                From = "de {0}",
                UpdateAvailable = "Mise à jour {0} disponible",
                Loading = "Téléchargement de la mise à jour…",
                UpdateBadge = "Mise à jour disponible",
                AddAttachmentsDialog = "Ajouter des pièces jointes",
            },
            ["es"] = new Loc
            {
                PaneTitle = "Notas de correo",
                InsertScreenshot = "Insertar captura",
                AttachFile = "Adjuntar archivo",
                Attachments = "Adjuntos",
                Open = "Abrir",
                Remove = "Quitar",
                AutoSaved = "guardado automáticamente",
                Saving = "Guardando…",
                SavedAt = "guardado {0}",
                NoMailSelected = "Ningún correo seleccionado",
                NoImageInClipboard = "No hay imagen en el portapapeles.",
                ErrorPrefix = "Error: {0}",
                NoteHeader = "Nota n.º {0}  {1}",
                NoSubject = "(sin asunto)",
                From = "de {0}",
                UpdateAvailable = "Actualización {0} disponible",
                Loading = "Descargando actualización…",
                UpdateBadge = "Actualización disponible",
                AddAttachmentsDialog = "Agregar adjuntos",
            },
            ["it"] = new Loc
            {
                PaneTitle = "Note e-mail",
                InsertScreenshot = "Inserisci screenshot",
                AttachFile = "Allega file",
                Attachments = "Allegati",
                Open = "Apri",
                Remove = "Rimuovi",
                AutoSaved = "salvato automaticamente",
                Saving = "Salvataggio…",
                SavedAt = "salvato {0}",
                NoMailSelected = "Nessuna e-mail selezionata",
                NoImageInClipboard = "Nessuna immagine negli appunti.",
                ErrorPrefix = "Errore: {0}",
                NoteHeader = "Nota #{0}  {1}",
                NoSubject = "(senza oggetto)",
                From = "da {0}",
                UpdateAvailable = "Aggiornamento {0} disponibile",
                Loading = "Download dell'aggiornamento…",
                UpdateBadge = "Aggiornamento disponibile",
                AddAttachmentsDialog = "Aggiungi allegati",
            },
            ["pt"] = new Loc
            {
                PaneTitle = "Notas de e-mail",
                InsertScreenshot = "Inserir captura",
                AttachFile = "Anexar arquivo",
                Attachments = "Anexos",
                Open = "Abrir",
                Remove = "Remover",
                AutoSaved = "salvo automaticamente",
                Saving = "Salvando…",
                SavedAt = "salvo {0}",
                NoMailSelected = "Nenhum e-mail selecionado",
                NoImageInClipboard = "Nenhuma imagem na área de transferência.",
                ErrorPrefix = "Erro: {0}",
                NoteHeader = "Nota #{0}  {1}",
                NoSubject = "(sem assunto)",
                From = "de {0}",
                UpdateAvailable = "Atualização {0} disponível",
                Loading = "Baixando atualização…",
                UpdateBadge = "Atualização disponível",
                AddAttachmentsDialog = "Adicionar anexos",
            },
            ["nl"] = new Loc
            {
                PaneTitle = "E-mailnotities",
                InsertScreenshot = "Schermafbeelding invoegen",
                AttachFile = "Bestand bijvoegen",
                Attachments = "Bijlagen",
                Open = "Openen",
                Remove = "Verwijderen",
                AutoSaved = "automatisch opgeslagen",
                Saving = "Opslaan…",
                SavedAt = "opgeslagen {0}",
                NoMailSelected = "Geen e-mail geselecteerd",
                NoImageInClipboard = "Geen afbeelding op het klembord.",
                ErrorPrefix = "Fout: {0}",
                NoteHeader = "Notitie #{0}  {1}",
                NoSubject = "(geen onderwerp)",
                From = "van {0}",
                UpdateAvailable = "Update {0} beschikbaar",
                Loading = "Update downloaden…",
                UpdateBadge = "Update beschikbaar",
                AddAttachmentsDialog = "Bijlagen toevoegen",
            },
            ["pl"] = new Loc
            {
                PaneTitle = "Notatki e-mail",
                InsertScreenshot = "Wstaw zrzut ekranu",
                AttachFile = "Załącz plik",
                Attachments = "Załączniki",
                Open = "Otwórz",
                Remove = "Usuń",
                AutoSaved = "zapisano automatycznie",
                Saving = "Zapisywanie…",
                SavedAt = "zapisano {0}",
                NoMailSelected = "Nie wybrano wiadomości",
                NoImageInClipboard = "Brak obrazu w schowku.",
                ErrorPrefix = "Błąd: {0}",
                NoteHeader = "Notatka #{0}  {1}",
                NoSubject = "(bez tematu)",
                From = "od {0}",
                UpdateAvailable = "Aktualizacja {0} dostępna",
                Loading = "Pobieranie aktualizacji…",
                UpdateBadge = "Dostępna aktualizacja",
                AddAttachmentsDialog = "Dodaj załączniki",
            },
            ["ru"] = new Loc
            {
                PaneTitle = "Заметки к письмам",
                InsertScreenshot = "Вставить снимок экрана",
                AttachFile = "Прикрепить файл",
                Attachments = "Вложения",
                Open = "Открыть",
                Remove = "Удалить",
                AutoSaved = "сохраняется автоматически",
                Saving = "Сохранение…",
                SavedAt = "сохранено {0}",
                NoMailSelected = "Письмо не выбрано",
                NoImageInClipboard = "В буфере обмена нет изображения.",
                ErrorPrefix = "Ошибка: {0}",
                NoteHeader = "Заметка №{0}  {1}",
                NoSubject = "(без темы)",
                From = "от {0}",
                UpdateAvailable = "Обновление {0} доступно",
                Loading = "Загрузка обновления…",
                UpdateBadge = "Доступно обновление",
                AddAttachmentsDialog = "Добавить вложения",
            },
            ["tr"] = new Loc
            {
                PaneTitle = "E-posta notları",
                InsertScreenshot = "Ekran görüntüsü ekle",
                AttachFile = "Dosya ekle",
                Attachments = "Ekler",
                Open = "Aç",
                Remove = "Kaldır",
                AutoSaved = "otomatik kaydedildi",
                Saving = "Kaydediliyor…",
                SavedAt = "kaydedildi {0}",
                NoMailSelected = "E-posta seçilmedi",
                NoImageInClipboard = "Panoda görüntü yok.",
                ErrorPrefix = "Hata: {0}",
                NoteHeader = "Not #{0}  {1}",
                NoSubject = "(konu yok)",
                From = "gönderen: {0}",
                UpdateAvailable = "{0} güncellemesi mevcut",
                Loading = "Güncelleme indiriliyor…",
                UpdateBadge = "Güncelleme mevcut",
                AddAttachmentsDialog = "Ek ekle",
            },
            ["ja"] = new Loc
            {
                PaneTitle = "メールのメモ",
                InsertScreenshot = "スクリーンショットを挿入",
                AttachFile = "ファイルを添付",
                Attachments = "添付ファイル",
                Open = "開く",
                Remove = "削除",
                AutoSaved = "自動保存されます",
                Saving = "保存中…",
                SavedAt = "保存しました {0}",
                NoMailSelected = "メールが選択されていません",
                NoImageInClipboard = "クリップボードに画像がありません。",
                ErrorPrefix = "エラー: {0}",
                NoteHeader = "メモ #{0}  {1}",
                NoSubject = "（件名なし）",
                From = "差出人: {0}",
                UpdateAvailable = "アップデート {0} が利用可能",
                Loading = "更新をダウンロード中…",
                UpdateBadge = "更新があります",
                AddAttachmentsDialog = "添付ファイルを追加",
            },
            ["ko"] = new Loc
            {
                PaneTitle = "이메일 메모",
                InsertScreenshot = "스크린샷 삽입",
                AttachFile = "파일 첨부",
                Attachments = "첨부 파일",
                Open = "열기",
                Remove = "제거",
                AutoSaved = "자동 저장됨",
                Saving = "저장 중…",
                SavedAt = "저장됨 {0}",
                NoMailSelected = "선택된 메일 없음",
                NoImageInClipboard = "클립보드에 이미지가 없습니다.",
                ErrorPrefix = "오류: {0}",
                NoteHeader = "메모 #{0}  {1}",
                NoSubject = "(제목 없음)",
                From = "보낸 사람: {0}",
                UpdateAvailable = "업데이트 {0} 사용 가능",
                Loading = "업데이트 다운로드 중…",
                UpdateBadge = "업데이트 사용 가능",
                AddAttachmentsDialog = "첨부 파일 추가",
            },
            ["zh"] = new Loc
            {
                PaneTitle = "邮件笔记",
                InsertScreenshot = "插入截图",
                AttachFile = "附加文件",
                Attachments = "附件",
                Open = "打开",
                Remove = "移除",
                AutoSaved = "已自动保存",
                Saving = "正在保存…",
                SavedAt = "已保存 {0}",
                NoMailSelected = "未选择邮件",
                NoImageInClipboard = "剪贴板中没有图像。",
                ErrorPrefix = "错误: {0}",
                NoteHeader = "笔记 #{0}  {1}",
                NoSubject = "（无主题）",
                From = "来自 {0}",
                UpdateAvailable = "更新 {0} 可用",
                Loading = "正在下载更新…",
                UpdateBadge = "有可用更新",
                AddAttachmentsDialog = "添加附件",
            },
        };
    }
}
