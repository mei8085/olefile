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


class TestMaliciousSectorChains(unittest.TestCase):
    """Regression tests for write mode against malicious FAT/MiniFAT chains.

    A corrupted sector chain must never cause an out-of-bounds write, an
    unbounded loop or a half-written file: write_stream has to reject the
    whole chain before touching the file.
    """

    def setUp(self):
        self.ole_file = u"tests/images/test-ole-file.doc"
        with open(self.ole_file, 'rb') as f:
            self.original = f.read()
        # a FAT-based stream and a MiniFAT-based stream present in the fixture:
        self.fat_stream = 'WordDocument'
        self.mini_stream = '\x01CompObj'

    def _ole_bytes(self):
        return bytearray(self.original)

    def _fat_location(self, data):
        # first FAT sector is referenced by the first DIFAT entry in the
        # header (offset 76); the fixture uses a single FAT sector and no DIFAT
        fat_sect = struct.unpack('<I', bytes(data[76:80]))[0]
        return self._sector_offset(fat_sect)

    def _sector_offset(self, sect):
        return 512 * (sect + 1)

    def _patch_fat(self, data, index, value):
        fat_off = self._fat_location(data)
        fat = list(struct.unpack('<128I', bytes(data[fat_off:fat_off + 512])))
        fat[index] = value
        data[fat_off:fat_off + 512] = struct.pack('<128I', *fat)

    def _patch_minifat(self, data, index, value):
        ole = olefile.OleFileIO(io.BytesIO(bytes(data)))
        minifat_off = self._sector_offset(ole.minifatsect)
        ole.close()
        minifat = list(struct.unpack('<128I',
                                     bytes(data[minifat_off:minifat_off + 512])))
        minifat[index] = value
        data[minifat_off:minifat_off + 512] = struct.pack('<128I', *minifat)

    def _patch_stream_start(self, data, stream_name, value):
        # set the isectStart field (offset 116) of a directory entry
        ole = olefile.OleFileIO(io.BytesIO(bytes(data)))
        sid = ole._find(stream_name)
        chain = []
        sect = ole.first_dir_sector
        while sect != olefile.ENDOFCHAIN:
            chain.append(sect)
            sect = ole.fat[sect]
        entries_per_sector = ole.sector_size // 128
        dir_off = self._sector_offset(chain[sid // entries_per_sector]) \
            + (sid % entries_per_sector) * 128
        ole.close()
        data[dir_off + 116:dir_off + 120] = struct.pack('<I', value)

    def _assert_rejected_without_write(self, data, stream_name, payload):
        before = bytes(data)
        buf = io.BytesIO(bytearray(data))
        with olefile.OleFileIO(buf, write_mode=True) as ole:
            with self.assertRaises(IOError):
                ole.write_stream(stream_name, payload)
        # the failed write must not leave any half-written byte behind:
        self.assertEqual(buf.getvalue(), before)
        # and must never extend the file:
        self.assertEqual(len(buf.getvalue()), len(self.original))

    def test_fat_chain_loop_rejected(self):
        'A loop in the FAT chain is rejected, no infinite loop'
        data = self._ole_bytes()
        # WordDocument occupies sectors 0..7: close the loop back to sector 0
        self._patch_fat(data, 7, 0)
        self._assert_rejected_without_write(
            data, self.fat_stream, b'A' * 4096)

    def test_fat_chain_out_of_range_rejected(self):
        'A FAT pointer outside the table is rejected before any write'
        data = self._ole_bytes()
        self._patch_fat(data, 7, 0x00010000)
        self._assert_rejected_without_write(
            data, self.fat_stream, b'A' * 4096)

    def test_fat_reserved_pointer_rejected(self):
        'Reserved FAT values (FREESECT/FATSECT/DIFSECT) are rejected'
        for bad in (olefile.FREESECT, olefile.FATSECT, olefile.DIFSECT):
            data = self._ole_bytes()
            self._patch_fat(data, 7, bad)
            self._assert_rejected_without_write(
                data, self.fat_stream, b'A' * 4096)

    def test_fat_invalid_first_sector_rejected(self):
        'An invalid isectStart cannot trigger a seek/write outside the file'
        data = self._ole_bytes()
        self._patch_stream_start(data, self.fat_stream, olefile.FREESECT)
        self._assert_rejected_without_write(
            data, self.fat_stream, b'A' * 4096)

    def test_minifat_chain_loop_rejected(self):
        'A loop in the MiniFAT chain is rejected, no infinite loop'
        data = self._ole_bytes()
        # CompObj ministream uses mini sectors 0 -> 1 -> ENDOFCHAIN
        self._patch_minifat(data, 1, 0)
        self._assert_rejected_without_write(
            data, self.mini_stream, b'B' * 114)

    def test_minifat_chain_out_of_range_rejected(self):
        'A MiniFAT pointer outside the table is rejected before any write'
        data = self._ole_bytes()
        self._patch_minifat(data, 1, 0x0000FFFF)
        self._assert_rejected_without_write(
            data, self.mini_stream, b'B' * 114)

    def test_write_sect_rejects_out_of_range_index(self):
        'write_sect refuses sector indexes outside the file'
        buf = io.BytesIO(bytearray(self.original))
        with olefile.OleFileIO(buf, write_mode=True) as ole:
            for bad in (ole.nb_sect, ole.nb_sect + 10,
                        olefile.FREESECT, -1):
                with self.assertRaises(IOError):
                    ole.write_sect(bad, b'X' * ole.sector_size)
        # nothing was written at all
        self.assertEqual(buf.getvalue(), self.original)

    def test_legitimate_fat_and_mini_writes_still_work(self):
        'Valid compound documents can still be written back correctly'
        buf = io.BytesIO(bytearray(self.original))
        with olefile.OleFileIO(buf, write_mode=True) as ole:
            # FAT-based stream:
            fat_size = ole.get_size(self.fat_stream)
            fat_payload = bytes((i * 7 + 3) & 0xFF for i in range(fat_size))
            ole.write_stream(self.fat_stream, fat_payload)
            self.assertEqual(
                ole.openstream(self.fat_stream).read(), fat_payload)
            # MiniFAT-based stream:
            mini_size = ole.get_size(self.mini_stream)
            mini_payload = bytes((i * 5 + 1) & 0xFF for i in range(mini_size))
            ole.write_stream(self.mini_stream, mini_payload)
            self.assertEqual(
                ole.openstream(self.mini_stream).read(), mini_payload)
        # the file keeps the same size and stays a valid OLE document:
        self.assertEqual(len(buf.getvalue()), len(self.original))
        with olefile.OleFileIO(io.BytesIO(buf.getvalue())) as ole:
            self.assertEqual(
                ole.openstream(self.fat_stream).read(), fat_payload)
            self.assertEqual(
                ole.openstream(self.mini_stream).read(), mini_payload)

    def test_failed_disk_write_leaves_file_untouched(self):
        'A rejected write on a real file leaves it byte-identical'
        ole_file_copy = "tests/images/malicious-chain-copy.doc"
        if os.path.isfile(ole_file_copy):
            os.remove(ole_file_copy)
        try:
            data = self._ole_bytes()
            self._patch_fat(data, 7, 0)  # loop
            with open(ole_file_copy, 'wb') as f:
                f.write(bytes(data))
            with olefile.OleFileIO(ole_file_copy, write_mode=True) as ole:
                with self.assertRaises(IOError):
                    ole.write_stream(self.fat_stream, b'A' * 4096)
            with open(ole_file_copy, 'rb') as f:
                self.assertEqual(f.read(), bytes(data))
        finally:
            if os.path.isfile(ole_file_copy):
                os.remove(ole_file_copy)


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
