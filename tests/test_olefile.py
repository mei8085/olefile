"""
Tests for olefile
"""

from __future__ import print_function
import unittest

import os, sys
import io
import struct
from shutil import copy2

# Insert the olefile package relative to this test dir in the sys.path
# to make sure we test the right one, in case another olefile version
# is also installed in the main python path.
_thismodule_dir = os.path.normpath(os.path.abspath(os.path.dirname(__file__)))
# print('_thismodule_dir = %r' % _thismodule_dir)
_parent_dir = os.path.normpath(os.path.join(_thismodule_dir, '..'))
# print('_parent_dir = %r' % _thirdparty_dir)
if _parent_dir not in sys.path:
    sys.path.insert(0, _parent_dir)

import olefile


class Test_isOleFile(unittest.TestCase):
    'Tests for the isOleFile function'

    def setUp(self):
        self.non_ole_file = u"tests/images/flower.jpg"
        self.ole_file = u"tests/images/test-ole-file.doc"

    def test_isOleFile_str_path_false(self):
        'Test isOleFile with a unicode filename of a non-OLE file'
        is_ole = olefile.isOleFile(filename=self.non_ole_file)
        self.assertFalse(is_ole)

    def test_isOleFile_str_path_true(self):
        'Test isOleFile with a unicode filename of an OLE file'
        is_ole = olefile.isOleFile(filename=self.ole_file)
        self.assertTrue(is_ole)

    def test_isOleFile_bytes_path_false(self):
        'Test isOleFile with a bytes filename of a non-OLE file'
        is_ole = olefile.isOleFile(filename=self.non_ole_file.encode('latin_1'))
        self.assertFalse(is_ole)

    def test_isOleFile_bytes_path_true(self):
        'Test isOleFile with a bytes filename of an OLE file'
        is_ole = olefile.isOleFile(filename=self.ole_file.encode('latin_1'))
        self.assertTrue(is_ole)

    def test_isOleFile_bytes_string_false(self):
        'Test isOleFile with contents of a non-OLE file in a bytes string as filename parameter'
        with open(self.non_ole_file, 'rb') as f:
            data = f.read()
        assert (len(data) >= olefile.MINIMAL_OLEFILE_SIZE)
        is_ole = olefile.isOleFile(filename=data)
        self.assertFalse(is_ole)

    def test_isOleFile_small_bytes_string_false(self):
        'Test isOleFile with contents of a small non-OLE file (<1536b) in a bytes string as filename parameter'
        data = b" " * 100
        # In this case, we expect a FileNotFoundError exception, because isOleFile treats filename as a path
        # due to issue #142
        if sys.version_info[0] < 3:
            # On Python 2.x, this case raises an IOError exception:
            expected_exception = IOError
        else:
            # On Python 3.x, this case raises a FileNotFoundError exception:
            expected_exception = FileNotFoundError
        with self.assertRaises(expected_exception):
            is_ole = olefile.isOleFile(filename=data)

    def test_isOleFile_bytes_string_true(self):
        'Test isOleFile with contents of an OLE file in a bytes string as filename parameter'
        with open(self.ole_file, 'rb') as f:
            data = f.read()
        assert(len(data) >= olefile.MINIMAL_OLEFILE_SIZE)
        is_ole = olefile.isOleFile(filename=data)
        self.assertTrue(is_ole)

    def test_isOleFile_bytes_data_false(self):
        'Test isOleFile with contents of a non-OLE file in a bytes string as data parameter'
        with open(self.non_ole_file, 'rb') as f:
            data = f.read()
        is_ole = olefile.isOleFile(data=data)
        self.assertFalse(is_ole)

    def test_isOleFile_bytes_data_true(self):
        'Test isOleFile with contents of an OLE file in a bytes string as data parameter'
        with open(self.ole_file, 'rb') as f:
            data = f.read()
        is_ole = olefile.isOleFile(data=data)
        self.assertTrue(is_ole)

    def test_isOleFile_small_bytes_data_false(self):
        'Test isOleFile with contents of a small non-OLE file (<1536b) in a bytes string as data parameter'
        data = b" " * 100
        is_ole = olefile.isOleFile(data=data)
        self.assertFalse(is_ole)

    def test_isOleFile_small_bytes_ole_magic_data_false(self):
        'Test isOleFile with contents of a small non-OLE file (<1536b) starting with OLE magic in a bytes string as data parameter'
        with open(self.ole_file, 'rb') as f:
            data = f.read(100)
        is_ole = olefile.isOleFile(data=data)
        self.assertFalse(is_ole)

    def test_isOleFile_file_false(self):
        'Test isOleFile with a non-OLE file object as filename parameter'
        with open(self.non_ole_file, 'rb') as f:
            is_ole = olefile.isOleFile(filename=f)
        self.assertFalse(is_ole)

    def test_isOleFile_file_true(self):
        'Test isOleFile with an OLE file object as filename parameter'
        with open(self.ole_file, 'rb') as f:
            is_ole = olefile.isOleFile(filename=f)
        self.assertTrue(is_ole)


class TestOleFileIO(unittest.TestCase):
    'Tests for the OleFileIO class'

    def setUp(self):
        self.non_ole_file = u"tests/images/flower.jpg"
        self.ole_file = u"tests/images/test-ole-file.doc"

    def test_context_manager(self):
        with olefile.OleFileIO(self.ole_file) as ole:
            exists = ole.exists('worddocument')
            self.assertTrue(exists)

    def test_exists_worddocument(self):
        ole = olefile.OleFileIO(self.ole_file)
        exists = ole.exists('worddocument')
        self.assertTrue(exists)
        ole.close()

    def test_exists_no_vba_macros(self):
        ole = olefile.OleFileIO(self.ole_file)
        exists = ole.exists('macros/vba')
        self.assertFalse(exists)
        ole.close()

    def test_get_type(self):
        ole = olefile.OleFileIO(self.ole_file)
        doc_type = ole.get_type('worddocument')
        self.assertEqual(doc_type, olefile.STGTY_STREAM)
        ole.close()

    def test_get_size(self):
        ole = olefile.OleFileIO(self.ole_file)
        size = ole.get_size('worddocument')
        self.assertGreater(size, 0)
        ole.close()

    def test_get_rootentry_name(self):
        ole = olefile.OleFileIO(self.ole_file)
        root = ole.get_rootentry_name()
        self.assertEqual(root, "Root Entry")
        ole.close()

    def test_meta(self):
        ole = olefile.OleFileIO(self.ole_file)
        meta = ole.get_metadata()
        self.assertEqual(meta.author, b"Laurence Ipsum")
        self.assertEqual(meta.num_pages, 1)
        ole.close()

    def test_minifat_writing(self):
        ole_file_copy = "tests/images/test-ole-file-copy.doc"
        minifat_stream_name = "\x01compobj"
        if os.path.isfile(ole_file_copy):
            os.remove(ole_file_copy)
        copy2(self.ole_file, ole_file_copy)

        ole = olefile.OleFileIO(ole_file_copy, write_mode = True)
        stream = ole.openstream(minifat_stream_name)
        self.assertTrue(stream.size < ole.minisectorcutoff)
        str_read = stream.read()
        self.assertTrue(len(str_read) == stream.size)
        self.assertTrue(str_read != b'\x00' * stream.size)
        stream.close()
        ole.write_stream(minifat_stream_name, b'\x00' * stream.size)
        ole.close()

        ole = olefile.OleFileIO(ole_file_copy)
        stream = ole.openstream(minifat_stream_name)
        self.assertTrue(stream.size < ole.minisectorcutoff)
        str_read_replaced = stream.read()
        self.assertTrue(len(str_read_replaced) == stream.size)
        self.assertTrue(str_read_replaced != str_read)
        self.assertTrue(str_read_replaced == b'\x00' * len(str_read))
        stream.close()
        ole.close()
        os.remove(ole_file_copy)

    def test_write_sect_too_large(self):
        'Test write_sect with data larger than a sector'
        # read OLE file in memory to avoid writing to disk
        with open(self.ole_file, 'rb') as f:
            file_in_memory = io.BytesIO(f.read())
        with olefile.OleFileIO(file_in_memory, write_mode=True) as ole:
            size = ole.sector_size + 1
            # Attempt to overwrite last sector (ole.nb_sect - 1) with data too large
            # it should raise ValueError:
            with self.assertRaises(ValueError):
                ole.write_sect(ole.nb_sect - 1, b"x" * size)


class TestPropertySetBounds(unittest.TestCase):
    """Count/length amplification bounds in property set parsing.

    A declared count must never be trusted on its own: the parser has to relate
    the property type, the vector element count, the element width and the
    bytes actually left before allocating or traversing anything.
    """

    def setUp(self):
        self.ole = olefile.OleFileIO()

    def parse_value(self, data, property_type):
        # offset 0 points just past the 4-byte type tag, as in getproperties()
        return self.ole._parse_property(data, 0, 2, property_type, False, [])

    # --- malicious counts are rejected before use ---------------------------

    def test_huge_vector_of_unsupported_type_rejected(self):
        with self.assertRaises(olefile.PropertySetError):
            self.parse_value(struct.pack('<I', 0xffffffff),
                             olefile.olefile.VT_VECTOR | olefile.olefile.VT_EMPTY)

    def test_huge_vector_of_fixed_width_type_rejected(self):
        for base in (olefile.olefile.VT_I4, olefile.olefile.VT_UI1,
                     olefile.olefile.VT_FILETIME):
            with self.assertRaises(olefile.PropertySetError):
                self.parse_value(struct.pack('<I', 0xffffffff),
                                 olefile.olefile.VT_VECTOR | base)

    def test_huge_variant_vector_rejected(self):
        with self.assertRaises(olefile.PropertySetError):
            self.parse_value(struct.pack('<I', 0xffffffff),
                             olefile.olefile.VT_VECTOR | olefile.olefile.VT_VARIANT)

    def test_huge_string_counts_rejected(self):
        for vt in (olefile.olefile.VT_LPSTR, olefile.olefile.VT_BSTR,
                   olefile.olefile.VT_BLOB, olefile.olefile.VT_CF):
            with self.assertRaises(olefile.PropertySetError):
                self.parse_value(struct.pack('<I', 0xffffffff) + b'x' * 8, vt)

    def test_huge_unicode_count_rejected(self):
        with self.assertRaises(olefile.PropertySetError):
            self.parse_value(struct.pack('<I', 0xffffffff) + b'ab',
                             olefile.olefile.VT_LPWSTR)

    def test_count_larger_than_remaining_bytes_rejected(self):
        # two VT_I4 elements need 8 bytes, but only 4 are present
        with self.assertRaises(olefile.PropertySetError):
            self.parse_value(struct.pack('<Ii', 2, 42),
                             olefile.olefile.VT_VECTOR | olefile.olefile.VT_I4)

    def test_truncated_string_payload_rejected(self):
        # length says 100 bytes, only 3 are present
        with self.assertRaises(olefile.PropertySetError):
            self.parse_value(struct.pack('<I', 100) + b'abc',
                             olefile.olefile.VT_LPSTR)

    def test_error_is_value_error_and_olefile_error(self):
        # the distinguishable error must remain catchable both ways
        self.assertTrue(issubclass(olefile.PropertySetError, ValueError))
        self.assertTrue(issubclass(olefile.PropertySetError, IOError))
        self.assertTrue(issubclass(olefile.PropertySetError, olefile.olefile.OleFileError))

    # --- legitimate values are unchanged ------------------------------------

    def test_valid_scalars_unchanged(self):
        VT = olefile.olefile
        self.assertEqual(self.parse_value(struct.pack('<h', -3), VT.VT_I2), -3)
        self.assertEqual(self.parse_value(struct.pack('<H', 0xffff), VT.VT_UI2), 65535)
        # VT_I4 is decoded as an unsigned 32-bit value by the original parser;
        # that encoding is preserved for compatibility (only VT_I2 is signed).
        self.assertEqual(self.parse_value(struct.pack('<i', -7), VT.VT_I4), 0xfffffff9)
        self.assertEqual(self.parse_value(struct.pack('<I', 4), VT.VT_UI4), 4)
        self.assertEqual(self.parse_value(struct.pack('<H', 0xffff), VT.VT_BOOL), True)
        self.assertEqual(self.parse_value(b'\x05', VT.VT_UI1), 5)
        self.assertEqual(self.parse_value(struct.pack('<I', 4) + b'abc\x00', VT.VT_LPSTR), b'abc')
        self.assertEqual(self.parse_value(struct.pack('<I', 3) + b'xyz', VT.VT_BLOB), b'xyz')

    def test_valid_unicode_string_unchanged(self):
        data = struct.pack('<I', 3) + u'Hi\x00'.encode('utf-16le')
        self.assertEqual(self.parse_value(data, olefile.olefile.VT_LPWSTR), u'Hi\x00')

    def test_unicode_string_length_read_from_value_header(self):
        # The Length field is the first 4 bytes of the value; the characters
        # follow. A following property must not leak into the decoded string.
        lpwstr = struct.pack('<I', 3) + u'Hi\x00'.encode('utf-16le')
        trailing = struct.pack('<I', 30) + b'SECRET\x00'   # next property bytes
        data = lpwstr + trailing
        value, size = self.ole._parse_property_basic(
            data, 0, 2, olefile.olefile.VT_LPWSTR, False, [])
        self.assertEqual(value, u'Hi\x00')
        self.assertEqual(size, len(lpwstr))

    def test_valid_vectors_unchanged(self):
        VT = olefile.olefile
        self.assertEqual(
            self.parse_value(struct.pack('<III', 2, 17, 23), VT.VT_VECTOR | VT.VT_UI4),
            [17, 23])
        self.assertEqual(self.parse_value(struct.pack('<I', 0), VT.VT_VECTOR | VT.VT_EMPTY), [])
        # a variant vector keeps its per-element type tags
        data = struct.pack('<I', 2) + struct.pack('<Ii', VT.VT_I4, 42) \
            + struct.pack('<IH', VT.VT_I2, 1)
        self.assertEqual(self.parse_value(data, VT.VT_VECTOR | VT.VT_VARIANT), [42, 1])

    def test_empty_properties_unchanged(self):
        VT = olefile.olefile
        self.assertIsNone(self.parse_value(b'', VT.VT_EMPTY))
        self.assertIsNone(self.parse_value(b'', VT.VT_NULL))

    # --- isolation: a failing property cannot poison the others ------------

    def build_stream(self, properties):
        """Build a minimal PropertySetStream from (pid, vt, payload) tuples."""
        count = len(properties)
        header_size = 8 + 8 * count
        body = b''
        offsets = []
        for pid, vt, payload in properties:
            offsets.append(header_size + len(body))
            body += struct.pack('<I', vt) + payload
        section = struct.pack('<II', header_size + len(body), count)
        for (pid, vt, payload), offset in zip(properties, offsets):
            section += struct.pack('<II', pid, offset)
        section += body
        stream = struct.pack('<HHI', 0xFFFE, 0, 0) + b'\x00' * 16 \
            + struct.pack('<I', 1)
        stream += b'\x00' * 16 + struct.pack('<I', 48) + section
        return stream

    def getproperties(self, properties, raise_defects=olefile.DEFECT_FATAL):
        data = self.build_stream(properties)
        ole = olefile.OleFileIO(raise_defects=raise_defects)
        ole.openstream = lambda name: io.BytesIO(data)
        return ole, ole.getproperties('props')

    def test_failed_property_does_not_affect_following_ones(self):
        VT = olefile.olefile
        properties = [
            # a malicious count that must be rejected
            (1, VT.VT_VECTOR | VT.VT_I4, struct.pack('<I', 0xffffffff)),
            # then perfectly valid values
            (2, VT.VT_LPSTR, struct.pack('<I', 3) + b'ok\x00'),
            (3, VT.VT_I4, struct.pack('<i', 7)),
        ]
        ole, props = self.getproperties(properties)
        self.assertEqual(props, {2: b'ok', 3: 7})
        self.assertTrue(ole.parsing_issues)

    def test_failed_property_is_not_published(self):
        VT = olefile.olefile
        ole, props = self.getproperties([
            (1, VT.VT_VECTOR | VT.VT_VARIANT, struct.pack('<I', 0xffffffff)),
        ])
        self.assertNotIn(1, props)

    def test_strict_mode_raises_distinguishable_error(self):
        VT = olefile.olefile
        with self.assertRaises(olefile.PropertySetError):
            self.getproperties(
                [(1, VT.VT_VECTOR | VT.VT_I4, struct.pack('<I', 0xffffffff))],
                raise_defects=olefile.DEFECT_INCORRECT)

    def test_section_size_underflow_is_rejected(self):
        # a section size below the 8-byte header used to make read() receive a
        # negative count (i.e. "read everything"); it must be rejected instead
        for size in (0, 1, 4, 7):
            data = struct.pack('<HHI', 0xFFFE, 0, 0) + b'\x00' * 16 \
                + struct.pack('<I', 1) + b'\x00' * 16 \
                + struct.pack('<I', 48) + struct.pack('<I', size) + b'\x00' * 32
            ole = olefile.OleFileIO()
            ole.openstream = lambda name, data=data: io.BytesIO(data)
            self.assertEqual(ole.getproperties('props'), {})
            self.assertTrue(ole.parsing_issues)


class FileHandleCloseTest(unittest.TestCase):
    """Test file handles are closed correctly."""

    def setUp(self):
        self.non_ole_file = "tests/images/flower.jpg"
        self.ole_file = "tests/images/test-ole-file.doc"

    def leaking_test_function(self):
        """Function that leaks an open OleFileIO."""
        ole = olefile.OleFileIO(self.ole_file)

    @unittest.skip('Cannot predict when __del__ is run, so cannot test that '
                   'it issues a warning')
    # requires python version 3.2 or higher
    def test_warning(self):
        """Test that warning is issued when ole file leaks open fp."""
        with self.assertWarns(olefile.OleFileIONotClosed):
            self.leaking_test_function()

    @unittest.skip('Cannot test attribute fp of OleFileIO instance that '
                   'failed to construct')
    def test_init_fail(self):
        """Test that file handle is closed if open() from __init__ fails."""
        ole = None
        try:
            ole = olefile.OleFileIO(self.non_ole_file)
            self.fail('Should have raised an exception')
        except Exception as exc:
            self.assertEqual(str(exc), 'not an OLE2 structured storage file')
            self.assertTrue(ole.fp.closed)   # ole is still None

    def test_context_manager(self):
        """Test that file handle is closed by context manager."""
        file_handle = None
        with olefile.OleFileIO(self.ole_file) as ole:
            file_handle = ole.fp
            self.assertFalse(file_handle.closed)
        self.assertTrue(file_handle.closed)

    def test_manual(self):
        """Test that simple manual close always closes self-created handle."""
        ole = olefile.OleFileIO(self.ole_file)
        self.assertFalse(ole.fp.closed)
        _ = ole.listdir()
        self.assertFalse(ole.fp.closed)
        ole.close()
        self.assertTrue(ole.fp.closed)

    # TODO: check if this test is still relevant
    # def test_fp_stays_open(self):
    #     """Test that fp is not automatically closed if provided by caller."""
    #     with open(self.ole_file, 'rb') as file_handle:
    #         self.assertFalse(file_handle.closed)
    #         with olefile.OleFileIO(file_handle) as ole:
    #             self.assertFalse(file_handle.closed)
    #             self.assertEqual(file_handle, ole.fp)
    #         # check that handle is still open, although ole is now closed
    #         self.assertFalse(file_handle.closed)
    #
    #         # do something with it
    #         file_handle.seek(0)
    #         self.assertTrue(olefile.isOleFile(file_handle))
    #
    #     # now should be closed
    #     self.assertTrue(file_handle.closed)


if __name__ == '__main__':
    unittest.main()
