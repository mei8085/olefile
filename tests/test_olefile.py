"""
Tests for olefile
"""

from __future__ import print_function
import unittest

import os, sys
import io
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


class TestForensicScan(unittest.TestCase):
    """Tests for the read-only forensic enumeration (OleFileIO.forensic_scan)"""

    # Layout of the synthetic OLE file (512-byte sectors):
    #   sector 0:     FAT
    #   sector 1:     directory
    #   sector 2:     MiniFAT
    #   sector 3:     MiniStream (root entry stream, 2 mini-sectors)
    #   sectors 4-12: stream "A" (4196 bytes, slack at end of sector 12)
    #   sector 13:    FREESECT, but contains data
    #   sector 14:    allocated (ENDOFCHAIN) but not referenced by any chain
    #   sectors 15-22: stream of orphan directory entry "ghost" (4096 bytes)
    # MiniStream: mini-sector 0 = stream "B" (20 bytes + slack),
    #             mini-sector 1 = FREESECT, but contains data
    NB_SECT = 23
    A_DATA = b'A' * 4196
    A_SLACK = b'\xCC' * 412
    B_DATA = b'B' * 20
    B_SLACK = b'\xAA' * 44
    FREE_MINI_DATA = b'\xBB' * 64
    FREE_SECT_DATA = b'\xDD' * 512
    UNREF_SECT_DATA = b'\xEE' * 512
    GHOST_DATA = b'G' * 4096

    @classmethod
    def build_ole_file(cls):
        """Build a synthetic OLE file in memory, with residual data."""
        import struct
        ENDOFCHAIN = 0xFFFFFFFE
        FREESECT = 0xFFFFFFFF
        FATSECT = 0xFFFFFFFD
        NOSTREAM = 0xFFFFFFFF

        # --- FAT (sector 0) ---
        fat = [FREESECT] * 128
        fat[0] = FATSECT
        fat[1] = ENDOFCHAIN                    # directory
        fat[2] = ENDOFCHAIN                    # MiniFAT
        fat[3] = ENDOFCHAIN                    # MiniStream
        for i in range(4, 12):                 # stream A chain
            fat[i] = i + 1
        fat[12] = ENDOFCHAIN
        fat[13] = FREESECT                     # free sector with data
        fat[14] = ENDOFCHAIN                   # allocated, unreferenced
        for i in range(15, 22):                # orphan stream chain
            fat[i] = i + 1
        fat[22] = ENDOFCHAIN
        sect_fat = struct.pack('<128L', *fat)

        # --- directory (sector 1) ---
        def dirent(name, etype, left, right, child, isect, size):
            name_raw = (name.encode('utf-16-le') + b'\x00\x00').ljust(64, b'\x00')
            return struct.pack('<64sHBBIII16sIQQIII', name_raw,
                               len(name) * 2 + 2, etype, 1,
                               left, right, child, b'\x00' * 16, 0, 0, 0,
                               isect, size, 0)
        sect_dir = b''.join([
            dirent('Root Entry', 5, NOSTREAM, NOSTREAM, 1, 3, 128),
            dirent('A', 2, NOSTREAM, 2, NOSTREAM, 4, len(cls.A_DATA)),
            dirent('B', 2, NOSTREAM, NOSTREAM, NOSTREAM, 0, len(cls.B_DATA)),
            dirent('ghost', 2, NOSTREAM, NOSTREAM, NOSTREAM, 15,
                   len(cls.GHOST_DATA)),
        ])

        # --- MiniFAT (sector 2): mini 0 used by B, mini 1 free ---
        minifat = [ENDOFCHAIN, FREESECT] + [FREESECT] * 126
        sect_minifat = struct.pack('<128L', *minifat)

        # --- MiniStream (sector 3) ---
        sect_ministream = (cls.B_DATA + cls.B_SLACK + cls.FREE_MINI_DATA
                           + b'\x00' * 384)

        # --- header ---
        header = struct.pack(
            '<8s16sHHHHHHLLLLLLLLLL',
            olefile.MAGIC, b'\x00' * 16, 0x003E, 3, 0xFFFE, 9, 6, 0, 0,
            0, 1, 1, 0, 0x1000, 2, 1, ENDOFCHAIN, 0)
        header += struct.pack('<109L', 0, *([FREESECT] * 108))

        sectors = [sect_fat, sect_dir, sect_minifat, sect_ministream]
        sectors += [b'A' * 512] * 8                       # sectors 4-11
        sectors.append(b'A' * 100 + cls.A_SLACK)          # sector 12
        sectors.append(cls.FREE_SECT_DATA)                # sector 13
        sectors.append(cls.UNREF_SECT_DATA)               # sector 14
        sectors += [b'G' * 512] * 8                       # sectors 15-22
        assert len(sectors) == cls.NB_SECT
        return header + b''.join(sectors)

    def setUp(self):
        self.ole_data = self.build_ole_file()

    def tearDown(self):
        if hasattr(self, 'ole') and self.ole is not None:
            self.ole.close()

    def open_ole(self):
        self.ole = olefile.OleFileIO(io.BytesIO(self.ole_data))
        return self.ole

    def test_normal_parsing_view(self):
        """The synthetic file must parse as a normal OLE file"""
        ole = self.open_ole()
        self.assertEqual(ole.parsing_issues, [])
        self.assertEqual(sorted(ole.listdir()), [['A'], ['B']])
        self.assertEqual(ole.openstream('A').read(), self.A_DATA)
        self.assertEqual(ole.openstream('B').read(), self.B_DATA)

    def test_forensic_scan_finds_residual_data(self):
        """forensic_scan must find all kinds of residual data"""
        ole = self.open_ole()
        items = ole.forensic_scan()
        by_type = {}
        for item in items:
            by_type.setdefault(item.item_type, []).append(item)

        # slack space at end of stream A (last FAT sector):
        slack = by_type[olefile.FORENSIC_STREAM_SLACK]
        self.assertEqual(len(slack), 1)
        self.assertEqual(slack[0].name, 'A')
        self.assertEqual(slack[0].path, ['A'])
        self.assertEqual(slack[0].sect, 12)
        self.assertEqual(slack[0].offset, 13 * 512 + 100)
        self.assertEqual(slack[0].size, 412)
        self.assertEqual(slack[0].data, self.A_SLACK)

        # slack space at end of ministream B (last mini-sector):
        mslack = by_type[olefile.FORENSIC_MINISTREAM_SLACK]
        self.assertEqual(len(mslack), 1)
        self.assertEqual(mslack[0].name, 'B')
        self.assertEqual(mslack[0].sect, 0)
        self.assertEqual(mslack[0].data, self.B_SLACK)

        # free mini-sector with data:
        freemini = by_type[olefile.FORENSIC_FREE_MINI_SECTOR]
        self.assertEqual(len(freemini), 1)
        self.assertEqual(freemini[0].sect, 1)
        self.assertEqual(freemini[0].data, self.FREE_MINI_DATA)

        # free sector with data:
        freesect = by_type[olefile.FORENSIC_FREE_SECTOR]
        self.assertEqual(len(freesect), 1)
        self.assertEqual(freesect[0].sect, 13)
        self.assertEqual(freesect[0].offset, 14 * 512)
        self.assertEqual(freesect[0].data, self.FREE_SECT_DATA)

        # allocated but unreferenced sector:
        unref = by_type[olefile.FORENSIC_UNREFERENCED_SECTOR]
        self.assertEqual(len(unref), 1)
        self.assertEqual(unref[0].sect, 14)
        self.assertEqual(unref[0].data, self.UNREF_SECT_DATA)

        # orphan directory entry, with original name and start sector:
        orphan = by_type[olefile.FORENSIC_ORPHAN_DIR_ENTRY]
        self.assertEqual(len(orphan), 1)
        self.assertEqual(orphan[0].name, 'ghost')
        self.assertEqual(orphan[0].sid, 3)
        self.assertEqual(orphan[0].sect, 15)
        self.assertEqual(orphan[0].size, len(self.GHOST_DATA))
        self.assertEqual(orphan[0].data, self.GHOST_DATA)
        self.assertEqual(orphan[0].entry_type, olefile.STGTY_STREAM)

        # exactly 6 items, nothing else:
        self.assertEqual(len(items), 6)

    def test_forensic_scan_does_not_report_metadata(self):
        """FAT, MiniFAT, directory and MiniStream sectors must not be
        reported as residual stream content"""
        ole = self.open_ole()
        items = ole.forensic_scan()
        for item in items:
            if item.item_type in (olefile.FORENSIC_FREE_SECTOR,
                                  olefile.FORENSIC_UNREFERENCED_SECTOR):
                # sectors 0 (FAT), 1 (directory), 2 (MiniFAT) and
                # 3 (MiniStream) are metadata, never residual content:
                self.assertNotIn(item.sect, (0, 1, 2, 3))

    def test_forensic_scan_does_not_alter_parsing_view(self):
        """forensic_scan must be isolated from the normal parsing view"""
        ole = self.open_ole()
        # baseline of the normal parsing view:
        listdir_before = ole.listdir()
        stream_a_before = ole.openstream('A').read()
        stream_b_before = ole.openstream('B').read()
        issues_before = list(ole.parsing_issues)
        # run the forensic scan:
        items = ole.forensic_scan()
        self.assertTrue(len(items) > 0)
        # the normal parsing view must be strictly unchanged:
        self.assertEqual(ole.listdir(), listdir_before)
        self.assertEqual(ole.openstream('A').read(), stream_a_before)
        self.assertEqual(ole.openstream('B').read(), stream_b_before)
        self.assertEqual(ole.parsing_issues, issues_before)
        # a second scan must return the same results:
        items2 = ole.forensic_scan()
        self.assertEqual(
            [(i.item_type, i.name, i.sid, i.sect, i.size) for i in items2],
            [(i.item_type, i.name, i.sid, i.sect, i.size) for i in items])

    def test_forensic_scan_on_real_file(self):
        """forensic_scan must run without error on a real OLE file"""
        ole = olefile.OleFileIO(u"tests/images/test-ole-file.doc")
        items = ole.forensic_scan()
        self.assertIsInstance(items, list)
        for item in items:
            self.assertIsInstance(item, olefile.OleForensicItem)
        # the normal parsing view still works after the scan:
        self.assertTrue(len(ole.listdir()) > 0)
        ole.close()


if __name__ == '__main__':
    unittest.main()
