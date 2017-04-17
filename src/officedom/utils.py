# -*- coding: utf-8 -*-

"""office DOM utilities"""

############################################################
#
# Copyright 2012, 2014, 2017 Mohammed El-Afifi
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
# file:         utils.py
#
# function:     office DOM helper utilities
#
# description:  contains helper utilities for office DOM API
#
# author:       Mohammed El-Afifi (ME)
#
# environment:  ActiveState Komodo IDE, version 6.1.3, build 66534,
#               windows xp professional
#
# notes:        This is a private program.
#
############################################################

class LightObject:

    """In-memory object

    This class represents a POJO object which holds only simple
    properties stored in memory but no indirect resources.
    """

    def __eq__(self, other):
        """Test if the two objects have the same content.

        `self` is this object.
        `other` is the other object.

        """
        # All public data attributes should be equal for the equality
        # test to succeed.
        attr_names = dir(self)

        for cur_attr in attr_names:
            if not cur_attr.startswith('_'):  # Exclude private attributes.

                attr_val = getattr(self, cur_attr)

                if not callable(attr_val) and \
                    attr_val != getattr(other, cur_attr):  # Exclude methods.
                    return False

        return True

    def __ne__(self, other):
        """Test if the two objects have different content.

        `self` is this object.
        `other` is the other object.

        """
        return not self == other


class _Wrapper:

    """Wrapper around a raw object"""

    def __init__(self, raw_obj):
        """Create a raw object wrapper.

        `self` is this wrapper.
        `raw_obj` is the raw object.

        """
        self._raw_obj = raw_obj


class ReadOnlyList(_Wrapper):

    """Read-only collection of objects"""

    def __init__(self, raw_list, conv_func):
        """Create a collection of objects.

        `self` is this collection of objects.
        `raw_list` is the list of COM objects.
        `conv_func` is a function to convert the raw list to a wrapper
                    list.

        """
        _Wrapper.__init__(self, raw_list)
        self._wrapper_list = map(conv_func, raw_list)

    def __getattr__(self, name):
        """Support immutable list operations.

        `self` is this collection of objects.
        `name` is the attribute.

        """
        if name in ["count", "index", "__len__", "__iter__", "__reversed__",
                    "__contains__"]:
            return getattr(self._wrapper_list, name)

        raise AttributeError()

    def __getitem__(self, key):
        """Support integer and string indices.

        `self` is this collection of objects.
        `key` is object index/key to look up.

        """
        try:
            return self._wrapper_list[key]  # integer indices
        except TypeError:  # string keys
            return self.get_wrapper(self._raw_obj(key))

    def get_wrapper(self, raw_obj):
        """Return the wrapper object for the given raw one.

        `self` is this collection of objects.
        `raw_obj` is the raw object to get whose wrapper.
        The method raises a ValueError if no object wraps the given raw
        one.

        """
        # Check if the object is wrapped.
        for cur_obj in self._wrapper_list:
            if cur_obj.raw_obj == raw_obj:
                return cur_obj

        raise ValueError()


class WrapperObject(_Wrapper, object):

    """Wrapper around a raw object"""

    def __init__(self, raw_obj):
        """Create a raw object wrapper.

        `self` is this wrapper.
        `raw_obj` is the raw object.

        """
        _Wrapper.__init__(self, raw_obj)

    @property
    def raw_obj(self):
        """Wrapped raw object

        `self` is this wrapper.

        """
        return self._raw_obj
