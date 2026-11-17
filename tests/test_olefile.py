"""
Tests for olefile
"""

from __future__ import print_function
import unittest

import os, sys
import io
import struct
import time
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


class TestSectorChainDefects(unittest.TestCase):
    """
    Tests for sector-chain hardening against FAT/MiniFAT loops (cycles) and
    cross-references between incompatible stream chains.
    All malformed documents are built in memory by mutating a valid OLE file.
    """

    def setUp(self):
        self.ole_file = u"tests/images/test-ole-file.doc"
        with open(self.ole_file, 'rb') as f:
            self.valid_data = f.read()
        # capture the legitimate stream contents and root chain, to prove that
        # a malformed chain neither reorders nor changes valid parsing:
        with olefile.OleFileIO(io.BytesIO(self.valid_data)) as ole:
            self.legit_streams = {
                tuple(path): ole.openstream(path).read()
                for path in ole.listdir()}
            ole.root.build_sect_chain(ole)
            self.legit_root_chain = list(ole.root.sect_chain)
            self.sector_size = ole.sector_size

    def _fat_entry_offset(self, fat_index):
        """Byte offset of a FAT entry in the OLE file (works with DIFAT too)."""
        with olefile.OleFileIO(io.BytesIO(self.valid_data)) as ole:
            entries_per_sector = ole.sector_size // 4
            fat_sect = ole.fat.index(olefile.FATSECT)
            fat_sect += fat_index // entries_per_sector
            index_in_sector = fat_index % entries_per_sector
        return (fat_sect + 1) * self.sector_size + index_in_sector * 4

    def _minifat_entry_offset(self, mini_index):
        """Byte offset of a MiniFAT entry inside the MiniFAT stream on disk."""
        with olefile.OleFileIO(io.BytesIO(self.valid_data)) as ole:
            minifat_start = ole.first_mini_fat_sector
            sector_size = ole.sector_size
            entries_per_sector = sector_size // 4
        fat_sect = minifat_start + mini_index // entries_per_sector
        index_in_sector = mini_index % entries_per_sector
        return (fat_sect + 1) * sector_size + index_in_sector * 4

    def _mutated(self, offsets_and_values):
        data = bytearray(self.valid_data)
        for offset, value in offsets_and_values:
            struct.pack_into('<I', data, offset, value)
        return io.BytesIO(bytes(data))

    def test_fat_loop_raises_defect(self):
        """A FAT chain loop must raise a defect instead of returning repeated data."""
        with olefile.OleFileIO(io.BytesIO(self.valid_data)) as ole:
            entry = ole.direntries[ole._find('WordDocument')]
            start = entry.isectStart
        # make the second sector point back at the first: start <-> start+1
        bio = self._mutated([(self._fat_entry_offset(start + 1), start)])
        with olefile.OleFileIO(bio, raise_defects=olefile.DEFECT_INCORRECT) as ole:
            with self.assertRaises(olefile.olefile.OleFileError) as ctx:
                ole.openstream('WordDocument')
            self.assertIn('loop', str(ctx.exception))

    def test_fat_loop_is_bounded_and_recorded_at_default_level(self):
        """At the default defect level the loop is recorded and traversal stops."""
        with olefile.OleFileIO(io.BytesIO(self.valid_data)) as ole:
            entry = ole.direntries[ole._find('WordDocument')]
            start, size = entry.isectStart, entry.size
        bio = self._mutated([(self._fat_entry_offset(start + 1), start)])
        t0 = time.time()
        with olefile.OleFileIO(bio) as ole:   # default: DEFECT_FATAL
            stream = ole.openstream('WordDocument')
            data = stream.read()
        self.assertLess(time.time() - t0, 2.0)
        self.assertTrue(any('loop' in msg for _, msg in ole.parsing_issues))
        # the chain was stopped before the loop, the stream is shorter than
        # declared but never contains repeated sectors:
        self.assertEqual(len(data), 2 * self.sector_size)
        self.assertLess(len(data), size)

    def test_fat_self_loop_does_not_hang_unknown_size_stream(self):
        """A self-referencing directory sector must not loop forever."""
        with olefile.OleFileIO(io.BytesIO(self.valid_data)) as ole:
            dir_start = ole.first_dir_sector
        bio = self._mutated([(self._fat_entry_offset(dir_start), dir_start)])
        t0 = time.time()
        try:
            olefile.OleFileIO(bio)
        except olefile.olefile.OleFileError:
            pass   # raising the defect is also acceptable
        self.assertLess(time.time() - t0, 2.0)

    def test_minifat_loop_raises_defect(self):
        """A loop in the MiniFAT chain must be detected like a FAT loop."""
        with olefile.OleFileIO(io.BytesIO(self.valid_data)) as ole:
            entry = ole.direntries[ole._find('\x01CompObj')]
            mini_start = entry.isectStart
        # point the first mini sector at itself:
        bio = self._mutated([(self._minifat_entry_offset(mini_start), mini_start)])
        with olefile.OleFileIO(bio, raise_defects=olefile.DEFECT_INCORRECT) as ole:
            with self.assertRaises(olefile.olefile.OleFileError) as ctx:
                ole.openstream('\x01CompObj')
            self.assertIn('loop', str(ctx.exception))

    def test_minifat_loop_bounded_and_recorded(self):
        bio = self._mutated([(self._minifat_entry_offset(0), 0)])
        t0 = time.time()
        with olefile.OleFileIO(bio) as ole:
            data = ole.openstream('\x01CompObj').read()
        self.assertLess(time.time() - t0, 2.0)
        self.assertTrue(any('loop' in msg for _, msg in ole.parsing_issues))
        self.assertLessEqual(len(data), 64)   # one mini sector at most

    def test_cross_referenced_fat_sector_raises_conflict(self):
        """Two streams sharing a FAT sector must produce an explicit conflict."""
        with olefile.OleFileIO(io.BytesIO(self.valid_data)) as ole:
            word = ole.direntries[ole._find('WordDocument')]
            table = ole.direntries[ole._find('1Table')]
        # repoint 1Table's first sector into the WordDocument chain:
        bio = self._mutated([
            (self._fat_entry_offset(table.isectStart), word.isectStart)])
        with olefile.OleFileIO(bio, raise_defects=olefile.DEFECT_INCORRECT) as ole:
            ole.openstream('WordDocument')          # claims the sectors first
            with self.assertRaises(olefile.olefile.OleFileError) as ctx:
                ole.openstream('1Table')            # re-enters a used sector
            self.assertIn('referenced by several stream chains',
                          str(ctx.exception))

    def test_cross_reference_does_not_affect_other_streams(self):
        """A conflicting chain is recorded but other objects parse unchanged."""
        with olefile.OleFileIO(io.BytesIO(self.valid_data)) as ole:
            word = ole.direntries[ole._find('WordDocument')]
            table = ole.direntries[ole._find('1Table')]
        bio = self._mutated([
            (self._fat_entry_offset(table.isectStart), word.isectStart)])
        with olefile.OleFileIO(bio) as ole:
            self.assertEqual(ole.openstream('WordDocument').read(),
                             self.legit_streams[('WordDocument',)])
            try:
                ole.openstream('1Table')
            except olefile.olefile.OleFileError:
                pass
            # unrelated streams are parsed exactly as in the valid file:
            self.assertEqual(
                ole.openstream('\x05SummaryInformation').read(),
                self.legit_streams[('\x05SummaryInformation',)])
            self.assertEqual(
                ole.openstream('\x05DocumentSummaryInformation').read(),
                self.legit_streams[('\x05DocumentSummaryInformation',)])
            self.assertTrue(any('referenced by several stream chains' in msg
                                for _, msg in ole.parsing_issues))

    def test_out_of_range_fat_pointer_raises(self):
        with olefile.OleFileIO(io.BytesIO(self.valid_data)) as ole:
            word = ole.direntries[ole._find('WordDocument')]
            start = word.isectStart
        bio = self._mutated([
            (self._fat_entry_offset(start + 1), 0x0ABCDEF0)])
        with olefile.OleFileIO(bio, raise_defects=olefile.DEFECT_INCORRECT) as ole:
            with self.assertRaises(olefile.olefile.OleFileError) as ctx:
                ole.openstream('WordDocument')
            self.assertIn('out of range', str(ctx.exception))

    def test_build_sect_chain_loop_is_bounded(self):
        """build_sect_chain (write path) must terminate on a cyclic chain."""
        with olefile.OleFileIO(io.BytesIO(self.valid_data)) as ole:
            table = ole.direntries[ole._find('1Table')]
            start = table.isectStart
        bio = self._mutated([(self._fat_entry_offset(start + 1), start)])
        t0 = time.time()
        with olefile.OleFileIO(bio, write_mode=True) as ole:
            entry = ole.direntries[ole._find('1Table')]
            entry.build_sect_chain(ole)
        self.assertLess(time.time() - t0, 2.0)
        # only the chain prefix before the loop is returned:
        self.assertEqual(entry.sect_chain, [start, start + 1])

    def test_build_sect_chain_out_of_range_stops(self):
        with olefile.OleFileIO(io.BytesIO(self.valid_data)) as ole:
            table = ole.direntries[ole._find('1Table')]
            start = table.isectStart
        bio = self._mutated([
            (self._fat_entry_offset(start + 1), 0x0ABCDEF0)])
        with olefile.OleFileIO(bio, write_mode=True) as ole:
            entry = ole.direntries[ole._find('1Table')]
            entry.build_sect_chain(ole)
        self.assertEqual(entry.sect_chain, [start, start + 1])

    def test_legitimate_chains_unchanged(self):
        """Valid chains keep their exact order and content, with no issues."""
        with olefile.OleFileIO(io.BytesIO(self.valid_data)) as ole:
            for path, expected in self.legit_streams.items():
                self.assertEqual(ole.openstream(list(path)).read(), expected)
            ole.root.build_sect_chain(ole)
            self.assertEqual(ole.root.sect_chain, self.legit_root_chain)
            self.assertEqual(ole.parsing_issues, [])


if __name__ == '__main__':
    unittest.main()
