# -*- coding: utf-8 -*-
"""
Word Count without Notes & Bibliography - a LibreOffice Writer extension.

  Tools > Word Count without Notes & Bibliography...
      Counts the words of the main text only.  Footnotes, endnotes, tables of contents and other
      indexes, headers/footers and anything tagged as bibliography are left out.

  Status bar
      The word count field at the bottom of the window shows this count instead of LibreOffice's
      own (selected text shows as "12 of 340 words").  Clicking it opens the detailed count.

  Count style
      The count window has a "Count style" choice: without notes and bibliography (the default),
      or everything, like LibreOffice's own count.  The choice is remembered and the status bar
      follows it.

  Language
      Text follows Tools > Options > Language Settings > Languages > Locale setting (language
      only, no dialects).  English is Canadian.  Words LibreOffice already translates (Words,
      Characters, Footnotes, Endnotes, Bibliography) are taken from LibreOffice's own
      translations (wc_labels.json); the rest is in wc_strings.json, which anyone can extend.

  Tools > Mark Selection as Bibliography (not counted)
  Tools > Remove Bibliography Mark
      Tags text (e.g. a Works Cited page) without changing how it looks.

The tag is a character style whose name ends in " (not counted)".  A plain tag style has no
formatting of its own.  If the text already uses a character style (say "Emphasis"), a twin style
"Emphasis (not counted)" is created that inherits everything from it, so nothing changes visually.
Paragraphs in LibreOffice's own "Bibliography 1" / "Bibliography Heading" styles, and generated
bibliography lists, are left out automatically.

Counting uses LibreOffice's own word counter (on a scratch document), so the numbers follow the
same rules as Tools > Word Count.
"""
import json
import os
import re
import threading
import time

try:
    import uno
    import unohelper
    from com.sun.star.beans import PropertyValue
except Exception:  # outside LibreOffice
    uno = None
    unohelper = None
    PropertyValue = None

TAG_SUFFIX = u" (not counted)"
BASE_TAG = u"Bibliography" + TAG_SUFFIX
BIB_PARA_STYLES = (u"Bibliography 1", u"Bibliography Heading", u"Bibliography Index")
NO_STYLE = (u"", u"Default Character Style", u"Standard")
COUNTED_PORTIONS = ("Text", "TextField", "ControlCharacter")


# ---------------------------------------------------------------------------
# language
# ---------------------------------------------------------------------------

EN = {
    "words": "Words", "chars": "Characters", "footnotes": "Footnotes", "endnotes": "Endnotes",
    "bibliography": "Bibliography / Works Cited",
    "counted_words": "Words counted", "counted_chars": "Characters counted (with spaces)",
    "left_out": "Left out:", "indexes_line": "Table of contents and other indexes",
    "compare": "For comparison, everything (Tools > Word Count, including headers/footers):",
    "compare2": "For comparison, without notes and bibliography:",
    "style_head": "Count style",
    "style_without": "Without notes and bibliography",
    "style_all": "Everything (like LibreOffice)",
    "click": "Click for details.",
    "tip_without": "Words counted without footnotes, endnotes, bibliography (Works Cited), indexes "
                   "and headers/footers.  Click for details.",
    "tip_all": "Everything, like LibreOffice's own count.  Click for details.",
    "indexes": "Indexes", "headfoot": "Headers/footers", "title": "Word Count", "document": "Document",
    "open_writer": "Open a Writer document first.",
    "select_first": "Select the text first - for a Works Cited page, put the cursor at the start of its "
                    "heading and press Ctrl+Shift+End.",
    "wrong": "Something went wrong:",
    "menu_count": "Word Count without Notes & Bibliography...",
    "chars_short": "characters",
}
LO_KEYS = ("words", "chars", "footnotes", "endnotes", "bibliography", "indexes", "headfoot", "title", "document")
_DIR_OVERRIDE = {"zh": "zh-CN", "no": "nb", "iw": "he", "fil": "tl", "sa": "sa-IN", "pa": "pa-IN",
                 "sw": "sw-TZ", "sat": "sat-Olck", "kmr": "kmr-Latn"}
_COMMA = ("fr", "de", "es", "it", "pt", "nl", "ru", "pl", "cs", "sv", "da", "nb", "nn", "fi", "tr",
          "hu", "ro", "uk", "el", "bg", "sk", "sl", "hr", "sr", "lt", "lv", "et", "ca", "gl", "eu",
          "id", "vi")


def _load_json(name):
    try:
        with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), name), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


class _Language(object):
    """Text in the language of Tools > Options > Language Settings > Languages > Locale setting.
    Language only, no dialects: English is always Canadian."""

    def __init__(self):
        self.lang = "en"
        self.country = ""
        self.lo = {}
        self.own = {}
        self.sep = ","
        self._loaded = False
        self._seen = 0.0

    def _config(self, ctx, path, name):
        cp = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.configuration.ConfigurationProvider", ctx)
        pv = PropertyValue()
        pv.Name = "nodepath"
        pv.Value = path
        acc = cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (pv,))
        return acc.getByName(name) or ""

    def update(self, ctx, force=False):
        now = time.time()
        if self._loaded and not force and now - self._seen < 3.0:
            return
        self._seen = now
        loc = ""
        try:
            loc = self._config(ctx, "/org.openoffice.Setup/L10N", "ooSetupSystemLocale")   # Locale setting
            if not loc:
                loc = self._config(ctx, "/org.openoffice.System/L10N", "Locale")          # "Default": the OS
            if not loc:
                loc = self._config(ctx, "/org.openoffice.Setup/L10N", "ooLocale")
        except Exception:
            pass
        parts = (loc or "en-US").replace("_", "-").split("-")
        lang = parts[0].lower() or "en"
        self.country = parts[1] if len(parts) > 1 else ""
        if not self._loaded:
            self.lo = _load_json("wc_labels.json")
            self.own = _load_json("wc_strings.json")
            self._loaded = True
        self.lang = lang
        folder = "zh-TW" if (lang == "zh" and self.country.upper() in ("TW", "HK", "MO")) \
            else _DIR_OVERRIDE.get(lang, lang)
        self.folder = folder
        self.lo_labels = dict((k, v.strip().rstrip(u": \uff1a")) for k, v in
                              (self.lo.get(folder) or self.lo.get(lang) or {}).items())
        self.own_labels = self.own.get(folder) or self.own.get(lang) or {}
        self.sep = ","
        try:
            from com.sun.star.lang import Locale
            ld = ctx.ServiceManager.createInstanceWithContext("com.sun.star.i18n.LocaleData", ctx)
            lc = Locale()
            lc.Language, lc.Country, lc.Variant = lang, self.country, ""
            self.sep = ld.getLocaleItem(lc).thousandSeparator or ","
        except Exception:
            self.sep = u" " if lang in _COMMA else ","

    @property
    def english(self):
        return self.lang == "en"

    def t(self, key):
        if self.english:
            return EN[key]
        if key in LO_KEYS:
            return self.lo_labels.get(key) or EN[key]
        if key == "counted_words":
            return self.lo_labels.get("words") or EN[key]
        if key == "counted_chars":
            return self.lo_labels.get("chars") or EN[key]
        if key == "indexes_line":
            return self.own_labels.get(key) or self.lo_labels.get("indexes") or EN[key]
        v = self.own_labels.get(key)
        if v is None:
            v = self._built(key)
        return EN[key] if v is None else v

    def _list(self, with_headfoot=False):
        """The left-out things, in LibreOffice's own words (whatever it has for this language)."""
        keys = ["footnotes", "endnotes", "bibliography", "indexes"] + (["headfoot"] if with_headfoot else [])
        return u", ".join(self.lo_labels[k] for k in keys if self.lo_labels.get(k))

    def _built(self, key):
        """Languages without my own sentences: build the text from LibreOffice's words and signs, so
        nothing English is mixed in."""
        lo = self.lo_labels
        doc = lo.get("document")
        lst = self._list()
        if key == "style_head":
            return lo.get("title")
        if key == "style_all":
            return (doc + u" (LibreOffice)") if doc else None
        if key == "style_without":
            return (doc + u" \u2212 " + lst) if doc and lst else None
        if key == "left_out":
            return u"\u2212"
        if key == "compare":
            return (doc + u" (LibreOffice):") if doc else None
        if key == "compare2":
            sw = self._built("style_without")
            return (sw + u":") if sw else None
        if key == "click":
            return u""
        if key == "tip_without":
            full = self._list(True)
            return (lo["words"] + u" \u2212 " + full) if lo.get("words") and full else None
        if key == "tip_all":
            return self._built("style_all")
        return None

    def num(self, n):
        return format(n, ",").replace(",", self.sep)


L = _Language()


def _lang(ctx=None, force=False):
    if ctx is None:
        ctx = uno.getComponentContext()
    try:
        L.update(ctx, force)
    except Exception:
        pass
    return L


# ---------------------------------------------------------------------------
# count style (remembered between sessions)
# ---------------------------------------------------------------------------

STYLE_WITHOUT, STYLE_ALL = "without", "all"
_STYLE = {"value": None}


def _settings_file(ctx):
    ps = ctx.ServiceManager.createInstanceWithContext("com.sun.star.util.PathSubstitution", ctx)
    return uno.fileUrlToSystemPath(ps.substituteVariables(u"$(user)/notes_wordcount.json", True))


def get_style(ctx=None):
    if _STYLE["value"] is None:
        value = STYLE_WITHOUT
        try:
            with open(_settings_file(ctx or uno.getComponentContext()), encoding="utf-8") as f:
                if json.load(f).get("style") == STYLE_ALL:
                    value = STYLE_ALL
        except Exception:
            pass
        _STYLE["value"] = value
    return _STYLE["value"]


def set_style(ctx, value):
    _STYLE["value"] = value
    try:
        with open(_settings_file(ctx), "w", encoding="utf-8") as f:
            json.dump({"style": value}, f)
    except Exception:
        pass



def _is_tag(name):
    return bool(name) and name.endswith(TAG_SUFFIX)


def _msg(ctx, frame, title, text, error=False):
    try:
        tk = ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
        from com.sun.star.awt.MessageBoxType import ERRORBOX, INFOBOX
        box = tk.createMessageBox(frame.ContainerWindow if frame is not None else None,
                                  ERRORBOX if error else INFOBOX, 1, title, text)
        box.execute()
    except Exception:
        pass


# ---------------------------------------------------------------------------
# collecting the text
# ---------------------------------------------------------------------------

class _Buckets(object):
    def __init__(self):
        self.body = []
        self.bib = []
        self.index = []      # tables of contents etc.
        self.notes = []      # footnotes
        self.ends = []       # endnotes


def _paragraph_kind(par):
    """'bib' | 'index' | None for a whole paragraph, from its style / the index it sits in."""
    try:
        if par.ParaStyleName in BIB_PARA_STYLES:
            return "bib"
    except Exception:
        pass
    try:
        idx = par.DocumentIndex
        if idx is not None:
            return "bib" if idx.supportsService("com.sun.star.text.Bibliography") else "index"
    except Exception:
        pass
    return None


def _portion_text(po):
    t = po.TextPortionType
    if t not in COUNTED_PORTIONS:
        return u""
    s = po.String or u""
    if t == "ControlCharacter" and not s.strip():
        return u" "
    return s


def _walk(text, buckets):
    """Sort every paragraph of an XText into body / bibliography / index text."""
    en = text.createEnumeration()
    while en.hasMoreElements():
        el = en.nextElement()
        try:
            if el.supportsService("com.sun.star.text.TextTable"):
                for name in el.CellNames:
                    _walk(el.getCellByName(name), buckets)
                continue
            if not el.supportsService("com.sun.star.text.Paragraph"):
                continue
        except Exception:
            continue
        kind = _paragraph_kind(el)
        body_parts, bib_parts = [], []
        pe = el.createEnumeration()
        while pe.hasMoreElements():
            po = pe.nextElement()
            s = _portion_text(po)
            if not s:
                continue
            tagged = False
            try:
                tagged = _is_tag(po.CharStyleName)
            except Exception:
                pass
            (bib_parts if tagged else body_parts).append(s)
        body_s, bib_s = u"".join(body_parts), u"".join(bib_parts)
        if kind == "bib":
            buckets.bib.append(body_s + u" " + bib_s)
        elif kind == "index":
            buckets.index.append(body_s + u" " + bib_s)
        else:
            if body_s.strip():
                buckets.body.append(body_s)
            if bib_s.strip():
                buckets.bib.append(bib_s)


def _scratch_count(ctx, pieces_by_key):
    """Count words/characters for several lists of paragraphs with LibreOffice's own counter."""
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    p = PropertyValue()
    p.Name = "Hidden"
    p.Value = True
    scratch = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, (p,))
    out = {}
    try:
        for key, pieces in pieces_by_key.items():
            scratch.Text.setString(u"\n".join(pieces))
            out[key] = (scratch.WordCount, scratch.CharacterCount)
    finally:
        scratch.close(True)
    return out


def _collect_body(doc):
    b = _Buckets()
    _walk(doc.Text, b)
    try:
        for i in range(doc.TextFrames.Count):
            _walk(doc.TextFrames.getByIndex(i), b)
    except Exception:
        pass
    return b


def count_words(doc, ctx):
    """Return a dict with the counts (see keys below)."""
    b = _collect_body(doc)
    try:
        for i in range(doc.Footnotes.Count):
            b.notes.append(doc.Footnotes.getByIndex(i).String)
    except Exception:
        pass
    try:
        for i in range(doc.Endnotes.Count):
            b.ends.append(doc.Endnotes.getByIndex(i).String)
    except Exception:
        pass
    counts = _scratch_count(ctx, {"body": b.body, "bib": b.bib, "index": b.index,
                                  "notes": b.notes, "ends": b.ends})
    return {
        "words": counts["body"][0], "characters": counts["body"][1],
        "bibliography": counts["bib"][0], "indexes": counts["index"][0],
        "footnotes": counts["notes"][0], "endnotes": counts["ends"][0],
        "libreoffice_total": doc.WordCount,
        "libreoffice_characters": doc.CharacterCount,
    }


def format_counts(c, style=STYLE_WITHOUT, ctx=None):
    lg = _lang(ctx)
    n = lg.num
    if style == STYLE_ALL:
        return u"\n".join([
            u"%s: %s" % (lg.t("style_head"), lg.t("style_all")),
            u"",
            u"%s: %s" % (lg.t("counted_words"), n(c["libreoffice_total"])),
            u"%s: %s" % (lg.t("counted_chars"), n(c["libreoffice_characters"])),
            u"",
            u"%s %s" % (lg.t("compare2"), n(c["words"]))])
    return u"\n".join([
        u"%s: %s" % (lg.t("style_head"), lg.t("style_without")),
        u"",
        u"%s: %s" % (lg.t("counted_words"), n(c["words"])),
        u"%s: %s" % (lg.t("counted_chars"), n(c["characters"])),
        u"",
        lg.t("left_out"),
    ] + [u"   %s: %s" % (lg.t(k2), n(c[k1])) for k1, k2 in
         (("footnotes", "footnotes"), ("endnotes", "endnotes"), ("bibliography", "bibliography"),
          ("indexes", "indexes_line")) if lg.english or k2 not in LO_KEYS + ("indexes_line",) or
         lg.lo_labels.get("indexes" if k2 == "indexes_line" else k2) or lg.own_labels.get(k2)] + [
        u"",
        u"%s %s" % (lg.t("compare"), n(c["libreoffice_total"]))])


# ---------------------------------------------------------------------------
# tagging
# ---------------------------------------------------------------------------

def _style_container(doc):
    return doc.StyleFamilies.getByName("CharacterStyles")


def _ensure_tag_style(doc, parent):
    """Name of a tag style that looks exactly like `parent` (a character style name or '')."""
    styles = _style_container(doc)
    name = BASE_TAG if parent in NO_STYLE else parent + TAG_SUFFIX
    if not styles.hasByName(name):
        st = doc.createInstance("com.sun.star.style.CharacterStyle")
        styles.insertByName(name, st)
        if parent not in NO_STYLE:
            st.ParentStyle = parent
    return name


def _portions_of(rng):
    """Text portions inside a selected range."""
    cur = rng.getText().createTextCursorByRange(rng)
    out = []
    en = cur.createEnumeration()
    while en.hasMoreElements():
        el = en.nextElement()
        try:
            if not el.supportsService("com.sun.star.text.Paragraph"):
                continue
            pe = el.createEnumeration()
            while pe.hasMoreElements():
                po = pe.nextElement()
                if po.TextPortionType == "Text" and (po.String or u""):
                    out.append(po)
        except Exception:
            continue
    return out


def _selected_ranges(controller):
    sel = controller.Selection
    ranges = []
    try:
        for i in range(sel.Count):
            r = sel.getByIndex(i)
            if r.supportsService("com.sun.star.text.TextRange"):
                ranges.append(r)
    except Exception:
        pass
    return [r for r in ranges if (r.String or u"")]


def mark_ranges(doc, ranges):
    """Tag the text of the given ranges. Returns the number of portions changed."""
    n = 0
    for rng in ranges:
        for po in _portions_of(rng):
            cur = po.CharStyleName
            if _is_tag(cur):
                continue
            po.CharStyleName = _ensure_tag_style(doc, cur)
            n += 1
    return n


def unmark_ranges(doc, ranges):
    n = 0
    styles = _style_container(doc)
    for rng in ranges:
        for po in _portions_of(rng):
            cur = po.CharStyleName
            if not _is_tag(cur):
                continue
            parent = u""
            try:
                parent = styles.getByName(cur).ParentStyle or u""
            except Exception:
                pass
            if parent:
                po.CharStyleName = parent
            else:
                po.setPropertyToDefault("CharStyleName")
            n += 1
    return n


# ---------------------------------------------------------------------------
# status bar
# ---------------------------------------------------------------------------

STATUS_COMMAND = u".uno:StateWordCount"

_STATUSBARS = []          # live status bar controllers (so tagging can refresh them)


def _refresh_all_statusbars():
    for c in list(_STATUSBARS):
        try:
            c.invalidate()
            c.refresh(force=True)
        except Exception:
            pass


def status_tip(ctx=None, style=None):
    lg = _lang(ctx)
    style = style or get_style(ctx)
    if lg.english:
        return EN["tip_all" if style == STYLE_ALL else "tip_without"]
    if style == STYLE_ALL:
        text = lg.t("style_all")
    else:
        full = lg._list(True)
        text = (lg.t("words") + u" \u2212 " + full) if full and lg.lo_labels.get("words") else lg.t("style_without")
    click = lg.t("click")
    return (text + u". " + click) if click else text


def format_status(total, selected=None, ctx=None):
    """'1,234 words, 7,890 characters' - or '12 of 1,234 words, ...' when text is selected."""
    lg = _lang(ctx)
    n = lg.num
    w, ch = total
    if lg.english:
        def noun(k, one, many):
            return one if k == 1 else many
        if selected is not None:
            sw, sc = selected
            return u"%s of %s %s, %s of %s %s" % (n(sw), n(w), noun(w, u"word", u"words"),
                                                  n(sc), n(ch), noun(ch, u"character", u"characters"))
        return u"%s %s, %s %s" % (n(w), noun(w, u"word", u"words"), n(ch), noun(ch, u"character", u"characters"))
    words = lg.t("words")
    chars = lg.own_labels.get("chars_short") or lg.t("chars")
    if selected is not None:
        sw, sc = selected
        return u"%s / %s %s, %s / %s %s" % (n(sw), n(w), words, n(sc), n(ch), chars)
    return u"%s %s, %s %s" % (n(w), words, n(ch), chars)


def _selection_body(controller):
    """Paragraph buckets for the selected text (empty buckets when nothing is selected)."""
    b = _Buckets()
    n = 0
    for rng in _selected_ranges(controller):
        cur = rng.getText().createTextCursorByRange(rng)
        _walk(cur, b)
        n += 1
    return b, n



# ---------------------------------------------------------------------------
# commands
# ---------------------------------------------------------------------------

def _title(ctx):
    return _lang(ctx).t("menu_count").rstrip(u". \u2026")


def show_count_window(ctx, frame, c):
    """The count window: the breakdown, and the count style choice (remembered). Returns the style."""
    style = get_style(ctx)
    lg = _lang(ctx)
    try:
        smgr = ctx.ServiceManager
        dm = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
        dm.PositionX, dm.PositionY, dm.Width, dm.Height = 100, 80, 270, 190
        dm.Title = _title(ctx)

        def add(kind, name, x, y, w, h, **props):
            m = dm.createInstance("com.sun.star.awt.UnoControl%sModel" % kind)
            m.Name = name
            m.PositionX, m.PositionY, m.Width, m.Height = x, y, w, h
            for k, v in props.items():
                setattr(m, k, v)
            dm.insertByName(name, m)
            return m
        add("FixedText", "text", 8, 6, 254, 100, Label=format_counts(c, style, ctx), MultiLine=True)
        add("FixedText", "head", 8, 112, 254, 10, Label=lg.t("style_head") + u":")
        add("RadioButton", "rb_without", 14, 126, 248, 10, Label=lg.t("style_without"),
            State=1 if style == STYLE_WITHOUT else 0, TabIndex=1)
        add("RadioButton", "rb_all", 14, 140, 248, 10, Label=lg.t("style_all"),
            State=1 if style == STYLE_ALL else 0, TabIndex=2)
        add("Button", "ok", 100, 168, 70, 14, PushButtonType=1, DefaultButton=True, TabIndex=3)
        dialog = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
        dialog.setModel(dm)
        toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
        dialog.createPeer(toolkit, frame.ContainerWindow if frame is not None else None)
        result = dialog.execute()
        chosen = STYLE_ALL if dm.getByName("rb_all").State == 1 else STYLE_WITHOUT
        dialog.dispose()
        if result == 1 and chosen != style:
            set_style(ctx, chosen)
            _refresh_all_statusbars()
            return chosen
    except Exception:
        _msg(ctx, frame, _title(ctx), format_counts(c, style, ctx))
    return style


def run_command(ctx, frame, command, quiet_result=None):
    lg = _lang(ctx)
    title = _title(ctx)
    model = frame.Controller.Model if frame is not None else None
    if model is None or not model.supportsService("com.sun.star.text.TextDocument"):
        _msg(ctx, frame, title, lg.t("open_writer"), True)
        return None
    try:
        if command == "Count":
            c = count_words(model, ctx)
            if quiet_result is None:
                show_count_window(ctx, frame, c)
            return c
        ranges = _selected_ranges(frame.Controller)
        if not ranges:
            _msg(ctx, frame, title, lg.t("select_first"), True)
            return None
        n = mark_ranges(model, ranges) if command == "Mark" else unmark_ranges(model, ranges)
        _refresh_all_statusbars()
        return n
    except Exception as e:
        import traceback
        _msg(ctx, frame, title, "%s\n%s\n\n%s" % (lg.t("wrong"), e, traceback.format_exc()[-500:]), True)
        return None


if uno is not None and unohelper is not None:
    from com.sun.star.frame import XDispatchProvider, XDispatch
    from com.sun.star.lang import XInitialization, XServiceInfo

    class NotesWordCountHandler(unohelper.Base, XServiceInfo, XDispatchProvider, XDispatch, XInitialization):
        IMPL_NAME = "org.notescount.Handler"
        SERVICES = ("com.sun.star.frame.ProtocolHandler",)
        PROTOCOL = "org.notescount.cmd:"

        def __init__(self, ctx):
            self.ctx = ctx
            self.frame = None

        def initialize(self, args):
            if args:
                self.frame = args[0]

        def getImplementationName(self):
            return self.IMPL_NAME

        def supportsService(self, name):
            return name in self.SERVICES

        def getSupportedServiceNames(self):
            return self.SERVICES

        def queryDispatch(self, url, target, flags):
            return self if url.Protocol == self.PROTOCOL else None

        def queryDispatches(self, requests):
            return tuple(self.queryDispatch(r.FeatureURL, r.FrameName, r.SearchFlags) for r in requests)

        def dispatch(self, url, args):
            if url.Path in ("Count", "Mark", "Unmark"):
                run_command(self.ctx, self.frame, url.Path)

        def addStatusListener(self, listener, url):
            pass

        def removeStatusListener(self, listener, url):
            pass

    g_ImplementationHelper = unohelper.ImplementationHelper()
    g_ImplementationHelper.addImplementation(NotesWordCountHandler, NotesWordCountHandler.IMPL_NAME,
                                             NotesWordCountHandler.SERVICES)

    from com.sun.star.frame import XStatusListener, XStatusbarController
    from com.sun.star.util import XUpdatable
    from com.sun.star.lang import XComponent
    from com.sun.star.awt import XCallback

    class _Later(unohelper.Base, XCallback):
        """Runs a function on LibreOffice's main thread (posted from a timer thread)."""
        def __init__(self, fn):
            self.fn = fn

        def notify(self, data):
            try:
                self.fn()
            except Exception:
                pass

    class NotesWordCountStatusbar(unohelper.Base, XServiceInfo, XInitialization, XStatusListener,
                                  XStatusbarController, XUpdatable, XComponent):
        """The word count field of Writer's status bar: shows the count without notes and
        bibliography, and opens the detailed count when clicked."""
        IMPL_NAME = "org.notescount.StatusbarController"
        SERVICES = ("com.sun.star.frame.StatusbarController",)
        MIN_INTERVAL = 0.6        # seconds between full recounts while typing

        def __init__(self, ctx, *args):
            # LibreOffice's status bar factory hands the arguments (Frame, CommandURL,
            # StatusbarItem, ...) to the constructor, not to initialize().
            self.ctx = ctx
            self.frame = None
            self.item = None
            self.command = STATUS_COMMAND
            self.dispatch = None
            self.url = None
            self.disposed = False
            self.listeners = []
            self.total = None          # (words, characters) of the whole document
            self.total_sig = None
            self.total_time = 0.0
            self.last_cost = 0.0
            self.timer = None
            self.later = _Later(self._timer_fired)
            self.clicked = 0.0
            if args:
                self.initialize(args)

        # --- XInitialization / XUpdatable -------------------------------------------------
        def initialize(self, args):
            for a in args or ():
                name = getattr(a, "Name", None)
                if name == "Frame":
                    self.frame = a.Value
                elif name == "CommandURL":
                    self.command = a.Value
                elif name == "StatusbarItem":
                    self.item = a.Value
            if self not in _STATUSBARS:
                _STATUSBARS.append(self)

        def update(self):
            self._bind()
            self.refresh(force=True)

        def _bind(self):
            if self.dispatch is not None or self.frame is None:
                return
            try:
                u = uno.createUnoStruct("com.sun.star.util.URL")
                u.Complete = self.command
                tr = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.util.URLTransformer", self.ctx)
                u = tr.parseStrict(u)[1]
                d = self.frame.queryDispatch(u, "", 0)
                if d is not None:
                    d.addStatusListener(self, u)
                    self.dispatch, self.url = d, u
            except Exception:
                pass

        # --- XStatusListener --------------------------------------------------------------
        def statusChanged(self, event):
            self.refresh()

        def disposing(self, source):
            self.dispatch = None

        # --- the count --------------------------------------------------------------------
        def _model(self):
            try:
                m = self.frame.Controller.Model
                if m is not None and m.supportsService("com.sun.star.text.TextDocument"):
                    return m
            except Exception:
                pass
            return None

        def invalidate(self):
            self.total = None
            self.total_sig = None

        def refresh(self, force=False):
            if self.disposed or self.item is None:
                return
            model = self._model()
            if model is None:
                return
            try:
                ctrl = self.frame.Controller
                style = get_style(self.ctx)
                sel_buckets, nsel = _selection_body(ctrl)
                if style == STYLE_ALL:
                    sel_all = sel_buckets.body + sel_buckets.bib + sel_buckets.index
                    has_all = nsel > 0 and bool(sel_all)
                    total = (model.WordCount, model.CharacterCount)
                    sel = _scratch_count(self.ctx, {"sel": sel_all})["sel"] if has_all else None
                    self._show(format_status(total, sel, self.ctx), status_tip(self.ctx, style))
                    return
                has_sel = nsel > 0 and bool(sel_buckets.body)
                sig = (model.WordCount, model.CharacterCount)
                now = time.time()
                need_total = (self.total is None or sig != self.total_sig or
                              (not has_sel and now - self.total_time > 5.0))
                if need_total and not force and self.total is not None:
                    wait = max(self.MIN_INTERVAL, 4 * self.last_cost) - (now - self.total_time)
                    if wait > 0:
                        self._schedule(wait)
                        return
                pieces = {"sel": sel_buckets.body}
                if need_total:
                    t0 = time.time()
                    pieces["total"] = _collect_body(model).body
                    counts = _scratch_count(self.ctx, pieces)
                    self.total = counts["total"]
                    self.total_sig = sig
                    self.total_time = time.time()
                    self.last_cost = self.total_time - t0
                elif has_sel:
                    counts = _scratch_count(self.ctx, pieces)
                else:
                    counts = {}
                self._show(format_status(self.total, counts["sel"] if has_sel else None, self.ctx),
                           status_tip(self.ctx, style))
            except Exception:
                pass

        def _show(self, text, tip):
            if self.item.Text != text:
                self.item.Text = text
            if self.item.HelpText != tip:
                self.item.HelpText = tip

        def _schedule(self, delay):
            if self.timer is not None:
                return

            def fire():
                self.timer = None
                try:
                    cb = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.AsyncCallback", self.ctx)
                    cb.addCallback(self.later, None)
                except Exception:
                    pass
            t = threading.Timer(delay, fire)
            t.daemon = True
            self.timer = t
            t.start()

        def _timer_fired(self):
            if not self.disposed:
                self.refresh(force=True)

        # --- XStatusbarController: clicking opens the details -----------------------------
        def _open(self):
            now = time.time()
            if now - self.clicked < 0.5:       # click + mouse-up of one press must not open two prompts
                return
            self.clicked = now
            run_command(self.ctx, self.frame, "Count")
            self.refresh(force=True)

        def click(self, pos):
            self._open()

        def doubleClick(self, pos):
            pass

        def mouseButtonDown(self, event):
            return False

        def mouseMove(self, event):
            return False

        def mouseButtonUp(self, event):
            return False

        def command(self, pos, cmd, mouse_event, data):
            pass

        def paint(self, graphics, rect, item_id, style):
            pass

        # --- XComponent -------------------------------------------------------------------
        def dispose(self):
            if self.disposed:
                return
            self.disposed = True
            try:
                if self.dispatch is not None and self.url is not None:
                    self.dispatch.removeStatusListener(self, self.url)
            except Exception:
                pass
            self.dispatch = None
            if self in _STATUSBARS:
                _STATUSBARS.remove(self)
            for l in list(self.listeners):
                try:
                    ev = uno.createUnoStruct("com.sun.star.lang.EventObject")
                    ev.Source = self
                    l.disposing(ev)
                except Exception:
                    pass

        def addEventListener(self, l):
            self.listeners.append(l)

        def removeEventListener(self, l):
            if l in self.listeners:
                self.listeners.remove(l)

        # --- XServiceInfo -----------------------------------------------------------------
        def getImplementationName(self):
            return self.IMPL_NAME

        def supportsService(self, name):
            return name in self.SERVICES

        def getSupportedServiceNames(self):
            return self.SERVICES

    g_ImplementationHelper.addImplementation(NotesWordCountStatusbar, NotesWordCountStatusbar.IMPL_NAME,
                                             NotesWordCountStatusbar.SERVICES)
