# -*- coding: utf-8 -*-

"""wraps word DOM"""

############################################################
#
# Copyright 2012, 2014 Mohammed El-Afifi
# This file is part of pyofficedom.
#
# pyofficedom is free software: you can redistribute it and/or modify
# it under the terms of the GNU Lesser General Public License as
# published by the Free Software Foundation, either version 3 of the
# License, or (at your option) any later version.
#
# pyofficedom is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with pyofficedom.  If not, see
# <http://www.gnu.org/licenses/>.
#
# program:      python office DOM
#
# file:         word.py
#
# function:     word DOM API
#
# description:  contains all classes of word DOM API
#
# author:       Mohammed Safwat (MS)
#
# environment:  ActiveState Komodo IDE, version 6.1.3, build 66534,
#               windows xp professional
#
# notes:        This is a private program.
#
############################################################

from functools import partial
import itertools
from weakref import proxy

import pythoncom
import win32com.client

from utils import LightObject, ReadOnlyList, WrapperObject

NO_OBJ = "none"

class Application(object):

    """Word application"""

    def __init__(self):
        """Create a word application.

        `self` is this application.
        Hook to an active word application instance or start a new one
        if no current one is running.

        """
        appCls = "Word.Application"
        self._app = win32com.client.DispatchEx(appCls)
        self._docs = _Documents(self._app.Documents)
        self._langs = _Languages(self._app.Languages)
        self._templates = _Templates(self._app.Templates, self._docs)
        self._docs.tmpls = proxy(self._templates)
        self._normal_tmpl = self._templates.get_wrapper(
            self._app.NormalTemplate)

    # context manager support
    def __enter__(self):
        """Setup a context for this word application.

        `self` is this application.

        """
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        """Exit this word application.

        `self` is this application.
        `exc_type` is the type of raised exception if one was raised or
                   None otherwise.
        `exc_value` is the raised exception if one was raised or None
                    otherwise.
        `traceback` is the traceback when the exception occurred, if
                    any, or None otherwise.

        """
        self.quit()

    def quit(self, *args, **kwargs):
        """Close this word application.

        `self` is this application.
        Positional and keyword arguments are the same as those accepted
        by the corresponding method in word DOM API.

        """
        self._app.Quit(*args, **kwargs)

    @property
    def documents(self):
        """Collection of open documents

        `self` is this application.

        """
        return self._docs

    @property
    def languages(self):
        """Collection of languages

        `self` is this application.

        """
        return self._langs

    @property
    def normal_template(self):
        """Normal template

        `self` is this application.

        """
        return self._normal_tmpl

    @property
    def templates(self):
        """Collection of loaded templates

        `self` is this application.

        """
        return self._templates


class _Document(WrapperObject):

    """Word document

    This document class is stateful in the sense that it holds a
    permanent reference(throughout its lifetime) to the underlying
    document COM object.

    """

    def __init__(self, doc_list, doc):
        """Create a word document.

        `self` is this word document.
        `doc_list` is the document collection owning this document. This
                   list is notified upon closing this document.
        `doc` is the underlying COM object representing the document.

        """
        WrapperObject.__init__(self, doc)
        self.data = _LightDocument(doc)
        self._parent_docs = proxy(doc_list)

    # context manager support
    def __enter__(self):
        """Setup a context for this word document.

        `self` is this word document.

        """
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        """Close this word document.

        `self` is this word document.
        `exc_type` is the type of raised exception if one was raised or
                   None otherwise.
        `exc_value` is the raised exception if one was raised or None
                    otherwise.
        `traceback` is the traceback when the exception occurred, if
                    any, or None otherwise.

        """
        self.close()

    def close(self, *args, **kwargs):
        """Close this document.

        `self` is this word document.
        Positional and keyword arguments are the same as those accepted
        by the corresponding method in word DOM API.
        The method notifies the parent document list about closure.

        """
        self._raw_obj.Close(*args, **kwargs)
        self._parent_docs.remove(self)

    def save(self):
        """Save this document.

        `self` is this word document.
        The method synchronizes data to the underlying COM object before
        saving.

        """
        self.data.sync(self._raw_obj)
        self._raw_obj.Save()

    def save_as(self, *args, **kwargs):
        """Save this document to the given file.

        `self` is this word document.
        Positional and keyword arguments are the same as those accepted
        by the corresponding method in word DOM API.
        The method synchronizes data to the underlying COM object before
        saving.

        """
        self.data.sync(self._raw_obj)
        self._raw_obj.SaveAs(*args, **kwargs)

    @property
    def attached_template(self):
        """Reference template full name

        `self` is this word document.

        """
        return self.data.attached_template

    @attached_template.setter
    def attached_template(self, value):
        """Set the reference template for this document.

        `self` is this word document.
        `value` is the desired template, its name, or full name.
        The property notifies the parent document list about template
        change.

        """
        self._raw_obj.AttachedTemplate = str(value)
        self.data.attached_template = self.raw_obj.AttachedTemplate.FullName
        self._parent_docs.refresh_tmpls(self._raw_obj.AttachedTemplate)

    @property
    def name(self):
        """Lower case document name

        `self` is this word document.

        """
        return self._raw_obj.Name.lower()


class _Documents(ReadOnlyList):

    """Collection of documents"""

    def __init__(self, docs):
        """Create a collection of documents.

        `self` is this collection of documents.
        `docs` are the COM objects representing documents.

        """
        ReadOnlyList.__init__(self, docs, partial(_Document, self))
        self.tmpls = None

    def add(self, *args, **kwargs):
        """Add a new empty document.

        `self` is this collection of documents.
        Positional and keyword arguments are the same as those accepted
        by the corresponding method in word DOM API.
        The method returns the new document.
        If the new document references a template that isn't loaded, the
        template is loaded.

        """
        doc = self._raw_obj.Add(*args, **kwargs)
        self._load_tmpl(doc.AttachedTemplate)
        return self._add_new_doc(doc)

    def add_raw_doc(self, raw_doc):
        """Find/Add the raw document and return the wrapper one.

        `self` is this collection of documents.
        `raw_doc` is the raw document to find/add.
        If the document is already open, return its wrapper. Otherwise
        add it to this collection and return its wrapper. The method
        isn't intended for direct use by clients.

        """
        # Check if the document is already open.
        try:
            return self.get_wrapper(raw_doc)
        # This collection doesn't contain the given document, add it and
        # return its wrapper.
        except ValueError:
            return self._add_new_doc(raw_doc)

    def close(self, *args, **kwargs):
        """Close all documents.

        `self` is this collection of documents.
        Positional and keyword arguments are the same as those accepted
        by the corresponding method in word DOM API.
        The method updates the list of loaded templates as well.

        """
        self._raw_obj.Close(*args, **kwargs)
        self._wrapper_list = []
        self.tmpls.cleanup()

    def open(self, file_name, *args, **kwargs):
        """Open the given document file and return it.

        `self` is this collection of documents.
        `file_name` is the file to open.
        Positional and keyword arguments are the same as those accepted
        by the corresponding method in word DOM API.
        The method takes care of not opening the same document twice. If
        the requested document references a template that isn't loaded,
        the template is loaded.

        """
        doc = self._raw_obj.Open(file_name, *args, **kwargs)
        self._load_tmpl(doc.AttachedTemplate)
        return self.add_raw_doc(doc)

    def refresh_tmpls(self, new_tmpl):
        """Load the given template and purge stale ones.

        `self` is this collection of documents.
        `new_tmpl` is the raw template to load.
        The method isn't intended for direct use by clients.

        """
        self.tmpls.cleanup()
        self._load_tmpl(new_tmpl)

    def remove(self, doc):
        """Remove the given document from this collection.

        `self` is this collection of documents.
        `doc` is the document to remove.
        The method updates the list of loaded templates as well. The
        method isn't intended for direct use by clients.

        """
        self._wrapper_list.remove(doc)
        self.tmpls.cleanup()

    def save(self, *args, **kwargs):
        """Save all documents.

        `self` is this collection of documents.
        Positional and keyword arguments are the same as those accepted
        by the corresponding method in word DOM API.

        """
        self._raw_obj.Save(*args, **kwargs)

    def _add_new_doc(self, raw_doc):
        """Add a new raw document and return the wrapper one.

        `self` is this collection of documents.
        `raw_doc` is the new document to add.
        The method will always add a new wrapper for the given document
        even if one already exists.

        """
        self._wrapper_list.append(_Document(self, raw_doc))
        return self._wrapper_list[-1]

    def _load_tmpl(self, raw_tmpl):
        """Load the raw template(if necessary).

        `self` is this collection of documents.
        `raw_tmpl` is the raw template to load.

        """
        if self.tmpls:
            # Check if the template is already loaded.
            try:
                self.tmpls.get_wrapper(raw_tmpl)
            except ValueError:  # The template isn't loaded, load it.
                self.tmpls.add(_Template(self, raw_tmpl))


class _Language(WrapperObject):

    """language information"""

    def __init__(self, lang):
        """Create a language.

        `self` is this language.
        `lang` is the underlying COM object representing the language.

        """
        WrapperObject.__init__(self, lang)

    def __getattr__(self, name):
        """Support language properties.

        `self` is this language.
        `name` is the attribute.

        """
        prop_map = {"id": "ID", "name": "Name", "name_local": "NameLocal"}

        if name in prop_map:

            attr_val = getattr(self._raw_obj, prop_map[name])
            return attr_val.__class__(str(attr_val).lower())

        raise AttributeError()


class _Languages(ReadOnlyList):

    """Collection of languages"""

    def __init__(self, langs):
        """Create a collection of languages.

        `self` is this collection of languages.
        `docs` are the COM objects representing languages.

        """
        ReadOnlyList.__init__(self, langs, _Language)

    def Item(self, index):
        """Return the language at the specified index.

        `self` is this collection of languages.
        `index` is the language ID, name, or local name.

        """
        return self.get_wrapper(self._raw_obj(index))

    def get_wrapper(self, raw_obj):
        """Return the wrapper language for the given raw one.

        `self` is this collection of languages.
        `raw_obj` is the raw language to get whose wrapper.
        The method raises a ValueError if no language wraps the given
        raw one.

        """
        # Check if the language is wrapped.
        for cur_lang in self._wrapper_list:
            # Comparison is based on ID's.
            if cur_lang.raw_obj.ID == raw_obj.ID:
                return cur_lang

        raise ValueError()


class _LightDocument(LightObject, object):

    """Lightweight word document

    This document class is stateless; it doesn't hold a reference to the
    underlying document COM object.

    """

    def __init__(self, doc):
        """Create a lightweight word document.

        `self` is this word document.
        `doc` is the underlying COM object representing the document.

        """
        self._active_theme = doc.ActiveTheme
        self._tmpl = doc.AttachedTemplate.FullName
        self.active_writing_style = {}

        for lang in doc.Application.Languages:
            try:  # sorry, no clean way to know available languages
                entry_maker = partial(
                    self._style_entry, lang, doc.ActiveWritingStyle(lang.ID))
            except pythoncom.com_error:
                pass
            else:
                self.active_writing_style.update(
                    itertools.imap(entry_maker, ["ID", "Name", "NameLocal"]))

    def sync(self, doc):
        """Update the underlying COM object.

        `self` is this word document.
        `doc` is the underlying COM object to update.
        The method isn't intended for direct use by clients.

        """
        if self.active_theme == NO_OBJ:
            doc.RemoveTheme()
        else:
            doc.ApplyTheme(self._active_theme)

        # The active writing style property can't be updated.
        doc.AttachedTemplate = self._tmpl

    @property
    def active_theme(self):
        """Active theme

        `self` is this word document.

        """
        return self._active_theme.lower()  # lower case

    @active_theme.setter
    def active_theme(self, value):
        """Set the active theme for this document.

        `self` is this word document.
        `value` is the desired active theme.

        """
        self._active_theme = value

    @property
    def attached_template(self):
        """Reference template full name

        `self` is this word document.

        """
        return self._tmpl.lower()  # lower case

    @attached_template.setter
    def attached_template(self, value):
        """Set the reference template for this document.

        `self` is this word document.
        `value` is the desired template itself or its full name.

        """
        self._tmpl = str(value)

    @staticmethod
    def _style_entry(lang, style, attr):
        """Return a writing style tuple.

        `lang` is the language to work on.
        `style` is the active writing style of the language.
        `attr` is the key language attribute.
        The method creates and returns a tuple of the specified language
        attribute and its active writing style.

        """
        attr_val = getattr(lang, attr)
        return attr_val.__class__(str(attr_val).lower()), style.lower()


class _LightTemplate(LightObject):

    """Lightweight word template

    This template class is stateless; it doesn't hold a reference to the
    underlying template COM object.

    """

    def __init__(self, tmpl):
        """Create a lightweight word template.

        `self` is this word template.
        `tmpl` is the underlying COM object representing the template.

        """
        self.auto_text_entries = dict(
            (entry.Name, entry.Value) for entry in tmpl.AutoTextEntries)

    def sync(self, tmpl):
        """Update the underlying COM object.

        `self` is this word template.
        `tmpl` is the underlying COM object to update.

        """
        # Rewrite the autoText entries of the raw template.
        for entry in tmpl.AutoTextEntries:
            entry.Delete()

        entries = self.auto_text_entries.iteritems()

        for name, val in entries:
            tmpl.AutoTextEntries.Add(
                name, tmpl.Application.Selection.Range).Value = val


class _Template(WrapperObject):

    """Word template

    This template class is stateful in the sense that it holds a
    permanent reference(throughout its lifetime) to the underlying
    template COM object.

    """

    def __init__(self, docs, tmpl):
        """Create a word template.

        `self` is this word template.
        `docs` are the collection of open documents. This list is
               notified upon opening this template as a document.
        `tmpl` is the underlying COM object representing the template.

        """
        WrapperObject.__init__(self, tmpl)
        self.data = _LightTemplate(tmpl)
        self._docs = docs

    def __str__(self):
        """Return the full name of this template

        `self` is this word template.

        """
        return self._raw_obj.FullName

    def open_as_document(self):
        """Open this template as a document.

        `self` is this word template.

        """
        return self._docs.add_raw_doc(self._raw_obj.OpenAsDocument())

    def save(self):
        """Save this template.

        `self` is this word template.
        The template doesn't need to be previously opened as a document
        to be saved.

        """
        self.data.sync(self._raw_obj)
        self._raw_obj.Save()

    @property
    def full_name(self):
        """Full path to the template file

        `self` is this word template.

        """
        return self._raw_obj.FullName.lower()

    @property
    def name(self):
        """Lower case template name

        `self` is this word template.

        """
        return self._raw_obj.Name.lower()


class _Templates(ReadOnlyList):

    """Collection of templates"""

    def __init__(self, tmpls, docs):
        """Create a collection of templates.

        `self` is this collection of templates.
        `tmpls` are the COM objects representing templates.
        `docs` are the collection of open documents.

        """
        ReadOnlyList.__init__(self, tmpls, partial(_Template, docs))

    def add(self, tmpl):
        """Add the template.

        `self` is this collection of templates.
        `tmpl` is the template to add.
        The method isn't intended for direct use by clients.

        """
        self._wrapper_list.append(tmpl)

    def cleanup(self):
        """Remove unloaded templates.

        `self` is this collection of templates.
        The method purges any templates no longer referenced by any of
        the open documents. It isn't intended for direct use by clients.

        """
        count = 0
        num_of_tmpls = len(self._wrapper_list)

        while count < num_of_tmpls:
            # template still referenced
            if self._wrapper_list[count].raw_obj in self._raw_obj:
                count += 1
            else:  # template no longer referenced

                self._wrapper_list.remove(self._wrapper_list[count])
                num_of_tmpls -= 1
