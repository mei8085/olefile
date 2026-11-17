"""
Structured fuzz tests for olefile, focused on the cross-boundary between
sector chains (FAT / MiniFAT) and the directory tree.

Malformed compound documents are built by mutating a known-good in-memory
OLE file at well-defined structural slots (header fields, FAT entries,
directory entry fields, MiniFAT entries). The tests verify that:

1. Illegal sector references and directory cycles are rejected in a
   controlled way: parsing neither reads out of bounds nor falls into
   unbounded processing (only OleFileError / IOError may be raised).
2. Legitimate documents keep their directory semantics unchanged, and
   rejected samples retain enough structural information (non-empty error
   messages, parsing_issues entries, dumped sample + seed + mutation
   description) to locate the problem.
"""

from __future__ import print_function

import io
import os
import random
import struct
import sys
import tempfile
import unittest

# Insert the olefile package relative to this test dir in the sys.path
# to make sure we test the right one, in case another olefile version
# is also installed in the main python path.
_thismodule_dir = os.path.normpath(os.path.abspath(os.path.dirname(__file__)))
_parent_dir = os.path.normpath(os.path.join(_thismodule_dir, '..'))
if _parent_dir not in sys.path:
    sys.path.insert(0, _parent_dir)

import olefile
from olefile import OleFileIO
from olefile.olefile import OleFileError


# --- Constants ----------------------------------------------------------------

SECTOR_SIZE = 512
FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF

STGTY_STORAGE = 1
STGTY_STREAM = 2
STGTY_ROOT = 5

# Sector layout of the valid document built below:
SID_FAT = 0         # the single FAT sector
SID_DIR = 1         # the directory sector (4 entries)
SID_MINIFAT = 2     # the MiniFAT sector
SID_MINISTREAM = 3  # first sector of the ministream (root entry stream)
SID_BIG = 4         # first of 8 sectors holding BigStream (4..11)
NUM_SECTORS = 12

# Exceptions that count as a *controlled* rejection by the library.
# OleFileError subclasses IOError; the library also raises bare IOError
# (e.g. "file not found", "incorrect DIFAT"). Anything else (IndexError,
# struct.error, ValueError, RecursionError, ...) means the parser escaped
# its bounds and must fail the test.
CONTROLLED_EXCEPTIONS = (OleFileError, IOError)


# --- Helpers to build a minimal valid OLE compound document -------------------

def build_direntry(name, entry_type, sid_left=NOSTREAM, sid_right=NOSTREAM,
                   sid_child=NOSTREAM, isect_start=0, size=0):
    """Build one 128-bytes directory entry (see OleDirectoryEntry.STRUCT_DIRENTRY)."""
    name_utf16 = name.encode('utf-16-le') if name else b''
    name_raw = name_utf16[:62].ljust(64, b'\x00')
    namelength = (len(name_utf16) + 2) if name else 0
    return struct.pack(
        '<64sHBBIII16sIQQIII',
        name_raw, namelength, entry_type, 1,  # color: black
        sid_left, sid_right, sid_child,
        b'\x00' * 16,  # CLSID
        0,             # user flags
        0, 0,          # creation / modification timestamps
        isect_start,
        size & 0xFFFFFFFF, (size >> 32) & 0xFFFFFFFF)


# Content of the valid document's streams, kept as module constants so tests
# can check that stream data survives parsing unchanged.
SMALL_STREAM_DATA = b'S' * 192   # 3 mini-sectors, stored in the MiniFAT
BIG_STREAM_DATA = b''.join(bytes([i]) * SECTOR_SIZE for i in range(SID_BIG, NUM_SECTORS))


def build_valid_ole():
    """
    Build a minimal valid OLE2 file (512-byte sectors) as bytes.

    Layout: header, then sector 0 = FAT, 1 = directory, 2 = MiniFAT,
    3 = ministream (root entry stream), 4..11 = BigStream data.

    Directory tree:
        Root Entry (sid 0, child -> sid 1)
        +- Storage1  (sid 1, storage, right -> sid 3, child -> sid 2)
        |   +- SmallStream (sid 2, stream, 192 bytes in MiniFAT)
        +- BigStream (sid 3, stream, 4096 bytes in FAT sectors 4..11)
    """
    # --- FAT sector ---
    fat = [FREESECT] * (SECTOR_SIZE // 4)
    fat[SID_FAT] = FATSECT
    fat[SID_DIR] = ENDOFCHAIN
    fat[SID_MINIFAT] = ENDOFCHAIN
    fat[SID_MINISTREAM] = ENDOFCHAIN
    for i in range(SID_BIG, NUM_SECTORS - 1):
        fat[i] = i + 1
    fat[NUM_SECTORS - 1] = ENDOFCHAIN
    fat_sector = struct.pack('<128I', *fat)

    # --- directory sector ---
    root = build_direntry('Root Entry', STGTY_ROOT, sid_child=1,
                          isect_start=SID_MINISTREAM, size=SECTOR_SIZE)
    storage = build_direntry('Storage1', STGTY_STORAGE, sid_right=3, sid_child=2)
    small = build_direntry('SmallStream', STGTY_STREAM,
                           isect_start=0, size=len(SMALL_STREAM_DATA))
    big = build_direntry('BigStream', STGTY_STREAM,
                         isect_start=SID_BIG, size=len(BIG_STREAM_DATA))
    dir_sector = root + storage + small + big
    assert len(dir_sector) == SECTOR_SIZE

    # --- MiniFAT sector: mini-sectors 0->1->2->ENDOFCHAIN for SmallStream ---
    minifat = [FREESECT] * (SECTOR_SIZE // 4)
    minifat[0] = 1
    minifat[1] = 2
    minifat[2] = ENDOFCHAIN
    minifat_sector = struct.pack('<128I', *minifat)

    # --- ministream sector (root entry stream): SmallStream data + padding ---
    ministream_sector = SMALL_STREAM_DATA + b'\x00' * (SECTOR_SIZE - len(SMALL_STREAM_DATA))

    # --- header ---
    header = struct.pack(
        '<8s16sHHHHHHLLLLLLLLLL',
        olefile.MAGIC,      # signature
        b'\x00' * 16,       # CLSID (must be zero)
        0x003E,             # minor version
        3,                  # dll version: 512-byte sectors
        0xFFFE,             # byte order
        9,                  # sector shift -> 512
        6,                  # mini sector shift -> 64
        0,                  # reserved
        0,                  # reserved
        0,                  # num dir sectors (must be 0 for dll version 3)
        1,                  # num FAT sectors
        SID_DIR,            # first directory sector
        0,                  # transaction signature
        0x1000,             # mini stream cutoff size
        SID_MINIFAT,        # first MiniFAT sector
        1,                  # num MiniFAT sectors
        ENDOFCHAIN,         # first DIFAT sector
        0,                  # num DIFAT sectors
    )
    # DIFAT array in header: first FAT sector is sector 0, rest is free
    header += struct.pack('<109I', 0, *([FREESECT] * 108))
    assert len(header) == SECTOR_SIZE

    doc = header + fat_sector + dir_sector + minifat_sector + ministream_sector
    doc += BIG_STREAM_DATA
    assert len(doc) == SECTOR_SIZE * (1 + NUM_SECTORS)
    return doc


# --- Offsets of structural slots inside the built document --------------------

def sect_offset(sid):
    """File offset of the first byte of a sector."""
    return SECTOR_SIZE * (sid + 1)

# header field offsets
HDR_DLL_VERSION = 0x1A
HDR_SECTOR_SHIFT = 0x1E
HDR_MINI_SECTOR_SHIFT = 0x20
HDR_NUM_FAT_SECTORS = 0x2C
HDR_FIRST_DIR_SECTOR = 0x30
HDR_MINI_CUTOFF = 0x38
HDR_FIRST_MINIFAT_SECTOR = 0x3C
HDR_NUM_MINIFAT_SECTORS = 0x40
HDR_FIRST_DIFAT_SECTOR = 0x44
HDR_NUM_DIFAT_SECTORS = 0x48

# directory entry field offsets (relative to the start of a 128-bytes entry)
DE_NAMELENGTH = 64
DE_TYPE = 66
DE_SID_LEFT = 68
DE_SID_RIGHT = 72
DE_SID_CHILD = 76
DE_ISECTSTART = 116
DE_SIZELOW = 120
DE_SIZEHIGH = 124

# sids of the directory entries in the valid document
DIR_SID_ROOT = 0
DIR_SID_STORAGE1 = 1
DIR_SID_SMALL = 2
DIR_SID_BIG = 3


def dir_entry_offset(sid, field):
    """File offset of a field inside directory entry `sid`."""
    return sect_offset(SID_DIR) + sid * 128 + field


def fat_entry_offset(index):
    """File offset of FAT entry `index`."""
    return sect_offset(SID_FAT) + 4 * index


def minifat_entry_offset(index):
    """File offset of MiniFAT entry `index`."""
    return sect_offset(SID_MINIFAT) + 4 * index


def set_u32(doc, offset, value):
    doc[offset:offset + 4] = struct.pack('<I', value & 0xFFFFFFFF)


def set_u16(doc, offset, value):
    doc[offset:offset + 2] = struct.pack('<H', value & 0xFFFF)


def mutate_doc(**mutations):
    """
    Return a mutated copy of the valid document.

    mutations maps a description string to (offset, fmt, value) tuples,
    applied in order. fmt is one of 'u32', 'u16', 'u8'.
    """
    doc = bytearray(build_valid_ole())
    applied = []
    for desc, (offset, fmt, value) in mutations.items():
        if fmt == 'u32':
            set_u32(doc, offset, value)
        elif fmt == 'u16':
            set_u16(doc, offset, value)
        elif fmt == 'u8':
            doc[offset] = value & 0xFF
        else:
            raise ValueError('unknown fmt %r' % fmt)
        applied.append('%s @0x%X=%#x' % (desc, offset, value))
    return bytes(doc), '; '.join(applied)


# --- Shared assertions --------------------------------------------------------

class ControlledParsingMixin(object):
    """
    Helpers asserting that parsing a (possibly malformed) document stays
    under control: only OleFileError/IOError may be raised, and every
    rejection carries a non-empty message with structural information.
    """

    def exercise_doc(self, data):
        """
        Parse the document and walk its whole directory tree, reading every
        stream. Returns ('parsed', stats) or ('rejected', message).
        Any exception other than OleFileError/IOError propagates and means
        the parser went out of control.
        """
        try:
            ole = OleFileIO(io.BytesIO(data),
                            raise_defects=olefile.DEFECT_INCORRECT)
        except CONTROLLED_EXCEPTIONS as exc:
            return ('rejected', str(exc))
        try:
            streams = ole.listdir(streams=True, storages=False)
            # also walk storages to exercise the whole tree
            ole.listdir(streams=True, storages=True)
            total_bytes = 0
            for path in streams:
                stream = ole.openstream(path)
                total_bytes += len(stream.read())
            # tolerated defects must retain structural information:
            # (exception type, non-empty message) tuples
            for exc_type, message in ole.parsing_issues:
                self.assertIsInstance(message, str)
                self.assertTrue(len(message) > 0,
                                'parsing issue with empty message')
            stats = ('entries=%d streams=%d bytes=%d'
                     % (len(ole.direntries), len(streams), total_bytes))
            return ('parsed', stats)
        except CONTROLLED_EXCEPTIONS as exc:
            return ('rejected', str(exc))
        finally:
            ole.close()

    def assert_controlled(self, data, context):
        """
        Assert that parsing `data` is either fully successful or rejected
        with an informative, controlled exception. `context` describes the
        sample for failure messages.
        """
        outcome, info = self.exercise_doc(data)
        if outcome == 'rejected':
            self.assertTrue(len(info) > 0,
                            'rejection without any message for %s' % context)
        return outcome, info

    def dump_failed_sample(self, data, tag):
        """
        Write a failing sample to a file so the structure can be inspected
        to locate the problem. Returns the file path.
        """
        fd, path = tempfile.mkstemp(prefix='olefuzz_%s_' % tag,
                                    suffix='.bin')
        with os.fdopen(fd, 'wb') as f:
            f.write(data)
        return path


# --- 1. Targeted malformed cases: illegal sector references -------------------

class TestIllegalSectorReferences(unittest.TestCase, ControlledParsingMixin):
    """
    Malformed sector chains must be rejected in a controlled way, either at
    parse time or when the affected stream is opened.
    """

    def parse(self, doc):
        return OleFileIO(io.BytesIO(doc),
                         raise_defects=olefile.DEFECT_INCORRECT)

    def test_first_dir_sector_out_of_range(self):
        doc, desc = mutate_doc(
            first_dir_sector=(HDR_FIRST_DIR_SECTOR, 'u32', 500))
        with self.assertRaises(OleFileError) as cm:
            self.parse(doc)
        self.assertIn('out of range', str(cm.exception))

    def test_first_dir_sector_is_endofchain(self):
        # ENDOFCHAIN as first directory sector: the directory stream is
        # empty, so no root entry can be loaded -> fatal defect
        doc, desc = mutate_doc(
            first_dir_sector=(HDR_FIRST_DIR_SECTOR, 'u32', ENDOFCHAIN))
        with self.assertRaises(OleFileError) as cm:
            self.parse(doc)
        self.assertIn('out of range', str(cm.exception))

    def test_stream_start_sector_out_of_range(self):
        # parsing succeeds, but opening the stream must be rejected
        doc, desc = mutate_doc(
            big_isectstart=(dir_entry_offset(DIR_SID_BIG, DE_ISECTSTART), 'u32', 500))
        ole = self.parse(doc)
        with self.assertRaises(OleFileError) as cm:
            ole.openstream('BigStream')
        self.assertIn('out of range', str(cm.exception))
        ole.close()

    def test_stream_size_larger_than_fat(self):
        # declared stream size needs more sectors than the FAT holds
        doc, desc = mutate_doc(
            big_size=(dir_entry_offset(DIR_SID_BIG, DE_SIZELOW), 'u32', 0xFFFFFFFF))
        ole = self.parse(doc)
        with self.assertRaises(OleFileError) as cm:
            ole.openstream('BigStream')
        self.assertIn('too large', str(cm.exception))
        ole.close()

    def test_stream_points_at_directory_sector(self):
        # cross-boundary: BigStream reuses the directory stream's sector,
        # which is already registered -> duplicate stream reference
        doc, desc = mutate_doc(
            big_isectstart=(dir_entry_offset(DIR_SID_BIG, DE_ISECTSTART), 'u32', SID_DIR))
        with self.assertRaises(OleFileError) as cm:
            self.parse(doc)
        self.assertIn('referenced twice', str(cm.exception))

    def test_ministream_points_at_minifat_sector(self):
        # cross-boundary: root entry stream reuses the MiniFAT sector
        doc, desc = mutate_doc(
            root_isectstart=(dir_entry_offset(DIR_SID_ROOT, DE_ISECTSTART), 'u32', SID_MINIFAT))
        with self.assertRaises(OleFileError) as cm:
            self.parse(doc)
        self.assertIn('referenced twice', str(cm.exception))

    def test_minifat_start_sector_out_of_range(self):
        doc, desc = mutate_doc(
            first_minifat=(HDR_FIRST_MINIFAT_SECTOR, 'u32', 0x00FFFFFF))
        ole = self.parse(doc)
        with self.assertRaises(OleFileError) as cm:
            ole.openstream(['Storage1', 'SmallStream'])
        self.assertIn('out of range', str(cm.exception))
        ole.close()

    def test_minifat_index_out_of_range(self):
        # first mini-sector index of SmallStream beyond the MiniFAT
        doc, desc = mutate_doc(
            small_isectstart=(dir_entry_offset(DIR_SID_SMALL, DE_ISECTSTART), 'u32', 100))
        ole = self.parse(doc)
        with self.assertRaises(OleFileError) as cm:
            ole.openstream(['Storage1', 'SmallStream'])
        self.assertIn('out of range', str(cm.exception))
        ole.close()

    def test_minifat_chain_entry_out_of_range(self):
        # a MiniFAT entry points beyond the ministream
        doc, desc = mutate_doc(
            minifat0=(minifat_entry_offset(0), 'u32', 0x00FFFFFF))
        ole = self.parse(doc)
        with self.assertRaises(OleFileError) as cm:
            ole.openstream(['Storage1', 'SmallStream'])
        self.assertIn('out of range', str(cm.exception))
        ole.close()

    def test_fat_chain_loop_is_bounded(self):
        # FAT loop 4 -> 5 -> 4 -> ... inside the declared stream size:
        # reading stays bounded by the declared size (8 sectors)
        doc, desc = mutate_doc(
            fat5=(fat_entry_offset(5), 'u32', SID_BIG))
        ole = self.parse(doc)
        data = ole.openstream('BigStream').read()
        self.assertEqual(len(data), len(BIG_STREAM_DATA))
        ole.close()

    def test_directory_chain_loop_is_bounded(self):
        # the directory sector points to itself: the directory stream is
        # read at most len(fat) times, then parsing completes
        doc, desc = mutate_doc(
            fat_dir=(fat_entry_offset(SID_DIR), 'u32', SID_DIR))
        ole = self.parse(doc)
        # bounded: at most len(fat) * sectorsize / 128 entries
        self.assertLessEqual(len(ole.direntries), NUM_SECTORS * SECTOR_SIZE // 128)
        # the 4 real entries are still found, directory semantics preserved
        self.assertEqual(ole.listdir(),
                         [['BigStream'], ['Storage1', 'SmallStream']])
        ole.close()

    def test_minifat_chain_loop_is_bounded(self):
        # MiniFAT loop 0 -> 1 -> 0 -> ... bounded by declared stream size
        doc, desc = mutate_doc(
            minifat1=(minifat_entry_offset(1), 'u32', 0))
        ole = self.parse(doc)
        data = ole.openstream(['Storage1', 'SmallStream']).read()
        self.assertEqual(len(data), len(SMALL_STREAM_DATA))
        ole.close()

    def test_truncated_file(self):
        # file cut after sector 5: BigStream chain runs past the FAT
        doc = build_valid_ole()[:sect_offset(6)]
        ole = self.parse(doc)
        with self.assertRaises(OleFileError):
            ole.openstream('BigStream').read()
        ole.close()

    def test_truncated_before_minifat(self):
        # only FAT + directory sectors left: parse succeeds, but both
        # streams must be rejected in a controlled way when opened
        doc = build_valid_ole()[:sect_offset(2)]
        ole = self.parse(doc)
        with self.assertRaises(OleFileError):
            ole.openstream('BigStream')
        with self.assertRaises(OleFileError):
            ole.openstream(['Storage1', 'SmallStream'])
        ole.close()

    def test_truncated_inside_directory(self):
        # only the FAT sector left: directory cannot be loaded
        doc = build_valid_ole()[:sect_offset(1)]
        with self.assertRaises(OleFileError):
            self.parse(doc)

    def test_invalid_sector_shift(self):
        doc, desc = mutate_doc(sector_shift=(HDR_SECTOR_SHIFT, 'u16', 0x1F))
        with self.assertRaises(OleFileError) as cm:
            self.parse(doc)
        self.assertIn('sector_shift', str(cm.exception))

    def test_sector_size_dllversion_mismatch(self):
        # sector shift for 4K sectors with dll version 3
        doc, desc = mutate_doc(sector_shift=(HDR_SECTOR_SHIFT, 'u16', 12))
        with self.assertRaises(OleFileError) as cm:
            self.parse(doc)
        self.assertIn('sector_size', str(cm.exception))

    def test_difat_first_sector_out_of_range(self):
        # DIFAT declared but its first sector index is ENDOFCHAIN
        doc = bytearray(build_valid_ole())
        set_u32(doc, HDR_NUM_FAT_SECTORS, 113)
        set_u32(doc, HDR_NUM_DIFAT_SECTORS, 1)
        # first DIFAT sector stays ENDOFCHAIN -> fatal defect
        with self.assertRaises(OleFileError) as cm:
            self.parse(bytes(doc))
        self.assertIn('DIFAT', str(cm.exception))


# --- 2. Targeted malformed cases: directory tree cycles and bad links ---------

class TestDirectoryTreeDefects(unittest.TestCase, ControlledParsingMixin):
    """
    Cycles and invalid references in the directory red-black tree must be
    rejected in a controlled way (no infinite recursion, no out-of-bounds).
    """

    def parse(self, doc):
        return OleFileIO(io.BytesIO(doc),
                         raise_defects=olefile.DEFECT_INCORRECT)

    def test_child_sid_out_of_range(self):
        doc, desc = mutate_doc(
            root_child=(dir_entry_offset(DIR_SID_ROOT, DE_SID_CHILD), 'u32', 9999))
        with self.assertRaises(OleFileError) as cm:
            self.parse(doc)
        self.assertIn('out of range', str(cm.exception))

    def test_root_child_points_to_root(self):
        # root is its own child: self-cycle
        doc, desc = mutate_doc(
            root_child=(dir_entry_offset(DIR_SID_ROOT, DE_SID_CHILD), 'u32', DIR_SID_ROOT))
        with self.assertRaises(OleFileError) as cm:
            self.parse(doc)
        self.assertIn('double reference', str(cm.exception))

    def test_tree_cycle_between_entries(self):
        # SmallStream's left pointer links back to Storage1: cycle
        doc, desc = mutate_doc(
            small_left=(dir_entry_offset(DIR_SID_SMALL, DE_SID_LEFT), 'u32', DIR_SID_STORAGE1))
        with self.assertRaises(OleFileError) as cm:
            self.parse(doc)
        self.assertIn('double reference', str(cm.exception))

    def test_entry_referenced_by_two_parents(self):
        # cross-parent duplicate: BigStream becomes a storage whose child
        # is SmallStream, already a child of Storage1
        doc = bytearray(build_valid_ole())
        doc[dir_entry_offset(DIR_SID_BIG, DE_TYPE)] = STGTY_STORAGE
        set_u32(doc, dir_entry_offset(DIR_SID_BIG, DE_SID_CHILD), DIR_SID_SMALL)
        set_u32(doc, dir_entry_offset(DIR_SID_BIG, DE_SIZELOW), 0)
        with self.assertRaises(OleFileError) as cm:
            self.parse(bytes(doc))
        self.assertIn('double reference', str(cm.exception))

    def test_sibling_link_out_of_range(self):
        doc, desc = mutate_doc(
            storage_right=(dir_entry_offset(DIR_SID_STORAGE1, DE_SID_RIGHT), 'u32', 0xFFFFFF))
        with self.assertRaises(OleFileError) as cm:
            self.parse(doc)
        self.assertIn('out of range', str(cm.exception))

    def test_invalid_entry_type(self):
        doc, desc = mutate_doc(
            small_type=(dir_entry_offset(DIR_SID_SMALL, DE_TYPE), 'u8', 9))
        with self.assertRaises(OleFileError) as cm:
            self.parse(doc)
        self.assertIn('storage type', str(cm.exception))

    def test_root_entry_not_root(self):
        doc, desc = mutate_doc(
            root_type=(dir_entry_offset(DIR_SID_ROOT, DE_TYPE), 'u8', STGTY_STREAM))
        with self.assertRaises(OleFileError) as cm:
            self.parse(doc)
        self.assertIn('root entry', str(cm.exception))

    def test_duplicate_root_entry(self):
        doc, desc = mutate_doc(
            storage_type=(dir_entry_offset(DIR_SID_STORAGE1, DE_TYPE), 'u8', STGTY_ROOT))
        with self.assertRaises(OleFileError) as cm:
            self.parse(doc)
        self.assertIn('root entry', str(cm.exception))

    def test_name_length_overflow(self):
        doc, desc = mutate_doc(
            small_namelen=(dir_entry_offset(DIR_SID_SMALL, DE_NAMELENGTH), 'u16', 0xFFFF))
        with self.assertRaises(OleFileError) as cm:
            self.parse(doc)
        self.assertIn('name length', str(cm.exception))


# --- 3. Legitimate documents: directory semantics unchanged -------------------

class TestValidDocumentSemantics(unittest.TestCase, ControlledParsingMixin):

    EXPECTED_STREAMS = [['BigStream'], ['Storage1', 'SmallStream']]
    EXPECTED_ALL = [['BigStream'], ['Storage1'], ['Storage1', 'SmallStream']]

    def open_valid(self):
        return OleFileIO(io.BytesIO(build_valid_ole()),
                         raise_defects=olefile.DEFECT_INCORRECT)

    def test_valid_doc_parses_without_defects(self):
        ole = self.open_valid()
        self.assertEqual(ole.parsing_issues, [])
        ole.close()

    def test_valid_doc_directory_tree(self):
        ole = self.open_valid()
        self.assertEqual(ole.listdir(), self.EXPECTED_STREAMS)
        self.assertEqual(ole.listdir(storages=True), self.EXPECTED_ALL)
        self.assertEqual(ole.get_rootentry_name(), 'Root Entry')
        self.assertEqual(ole.get_type('Storage1'), olefile.STGTY_STORAGE)
        self.assertEqual(ole.get_type('BigStream'), olefile.STGTY_STREAM)
        self.assertTrue(ole.exists('storage1/smallstream'))  # case-insensitive
        self.assertEqual(ole.get_size('BigStream'), len(BIG_STREAM_DATA))
        self.assertEqual(ole.get_size(['Storage1', 'SmallStream']),
                         len(SMALL_STREAM_DATA))
        ole.close()

    def test_valid_doc_stream_data(self):
        ole = self.open_valid()
        self.assertEqual(ole.openstream('BigStream').read(), BIG_STREAM_DATA)
        self.assertEqual(ole.openstream(['Storage1', 'SmallStream']).read(),
                         SMALL_STREAM_DATA)
        ole.close()

    def test_data_only_mutation_keeps_directory_semantics(self):
        # mutating stream *data* (not structure) must not change the
        # directory semantics of the document
        doc = bytearray(build_valid_ole())
        for i in range(sect_offset(SID_BIG), len(doc)):
            doc[i] ^= 0xFF
        ole = OleFileIO(io.BytesIO(doc),
                        raise_defects=olefile.DEFECT_INCORRECT)
        self.assertEqual(ole.parsing_issues, [])
        self.assertEqual(ole.listdir(), self.EXPECTED_STREAMS)
        self.assertEqual(ole.listdir(storages=True), self.EXPECTED_ALL)
        self.assertEqual(ole.get_size('BigStream'), len(BIG_STREAM_DATA))
        # only the stream content changed
        self.assertNotEqual(ole.openstream('BigStream').read(), BIG_STREAM_DATA)
        ole.close()

    def test_real_sample_file_unchanged(self):
        # regression: the shipped sample keeps its known directory semantics
        ole = OleFileIO(os.path.join('tests', 'images', 'test-ole-file.doc'),
                        raise_defects=olefile.DEFECT_INCORRECT)
        self.assertEqual(ole.parsing_issues, [])
        self.assertEqual(ole.listdir(),
                         [['\x01CompObj'],
                          ['\x05DocumentSummaryInformation'],
                          ['\x05SummaryInformation'],
                          ['1Table'],
                          ['WordDocument']])
        ole.close()


# --- 4. Structured fuzz over sector-chain x directory-tree slots --------------

# Structural slots targeted by the fuzzer: (name, offset, fmt).
# They cover the header fields driving chain loading, every FAT entry,
# every directory entry field involved in tree/chain linking, and the
# MiniFAT entries: exactly the cross-boundary between sector chains and
# the directory tree.
def _build_fuzz_slots():
    slots = []
    for name, offset, fmt in [
            ('dll_version', HDR_DLL_VERSION, 'u16'),
            ('sector_shift', HDR_SECTOR_SHIFT, 'u16'),
            ('mini_sector_shift', HDR_MINI_SECTOR_SHIFT, 'u16'),
            ('num_fat_sectors', HDR_NUM_FAT_SECTORS, 'u32'),
            ('first_dir_sector', HDR_FIRST_DIR_SECTOR, 'u32'),
            ('mini_cutoff', HDR_MINI_CUTOFF, 'u32'),
            ('first_minifat_sector', HDR_FIRST_MINIFAT_SECTOR, 'u32'),
            ('num_minifat_sectors', HDR_NUM_MINIFAT_SECTORS, 'u32'),
            ('first_difat_sector', HDR_FIRST_DIFAT_SECTOR, 'u32'),
            ('num_difat_sectors', HDR_NUM_DIFAT_SECTORS, 'u32'),
    ]:
        slots.append((name, offset, fmt))
    for i in range(NUM_SECTORS):
        slots.append(('fat[%d]' % i, fat_entry_offset(i), 'u32'))
    for sid in (DIR_SID_ROOT, DIR_SID_STORAGE1, DIR_SID_SMALL, DIR_SID_BIG):
        for field, fmt in [(DE_NAMELENGTH, 'u16'), (DE_TYPE, 'u8'),
                           (DE_SID_LEFT, 'u32'), (DE_SID_RIGHT, 'u32'),
                           (DE_SID_CHILD, 'u32'), (DE_ISECTSTART, 'u32'),
                           (DE_SIZELOW, 'u32'), (DE_SIZEHIGH, 'u32')]:
            slots.append(('dir[%d].field%d' % (sid, field),
                          dir_entry_offset(sid, field), fmt))
    for i in range(8):
        slots.append(('minifat[%d]' % i, minifat_entry_offset(i), 'u32'))
    return slots

FUZZ_SLOTS = _build_fuzz_slots()

# Pool of values that stress chain/tree boundaries: valid indices, indices
# just past the end, and the special FAT/chain markers.
FUZZ_VALUE_POOL = [
    0, 1, 2, 3, 4,
    NUM_SECTORS - 1, NUM_SECTORS, NUM_SECTORS + 1,
    0x7FFFFFFF, 0x80000000,
    0xFFFFFFFA,  # MAXREGSECT
    FATSECT, FREESECT, ENDOFCHAIN, NOSTREAM,
    0xFFFFFFFF,
]

FUZZ_SEEDS = 300


class TestStructuredFuzz(unittest.TestCase, ControlledParsingMixin):

    def mutate(self, seed):
        """Deterministically build a mutated document for the given seed."""
        rng = random.Random(seed)
        doc = bytearray(build_valid_ole())
        descriptions = []
        for _ in range(rng.randint(1, 4)):
            name, offset, fmt = rng.choice(FUZZ_SLOTS)
            if rng.random() < 0.7:
                value = rng.choice(FUZZ_VALUE_POOL)
            else:
                value = rng.getrandbits(32)
            if fmt == 'u32':
                set_u32(doc, offset, value)
            elif fmt == 'u16':
                set_u16(doc, offset, value)
            else:
                doc[offset] = value & 0xFF
            descriptions.append('%s@0x%X=%#x' % (name, offset, value))
        return bytes(doc), '; '.join(descriptions)

    def test_structured_fuzz_stays_controlled(self):
        rejected = 0
        parsed = 0
        for seed in range(FUZZ_SEEDS):
            doc, desc = self.mutate(seed)
            with self.subTest(seed=seed, mutations=desc):
                try:
                    outcome, info = self.assert_controlled(
                        doc, 'seed=%d mutations=%s' % (seed, desc))
                except Exception as exc:
                    # uncontrolled behaviour: keep the sample so the
                    # problem can be located, then fail with full context
                    path = self.dump_failed_sample(doc, 'seed%d' % seed)
                    self.fail(
                        'uncontrolled parsing for seed=%d\n'
                        'mutations: %s\n'
                        'sample dumped to: %s\n'
                        'exception: %r' % (seed, desc, path, exc))
                if outcome == 'rejected':
                    rejected += 1
                else:
                    parsed += 1
        # sanity: the mutations must actually exercise both outcomes
        self.assertGreater(rejected, 0, 'no mutated sample was ever rejected')
        self.assertGreater(parsed, 0, 'no mutated sample ever parsed')

    def test_fuzz_is_deterministic(self):
        # same seed -> same document, so failures are reproducible
        for seed in (0, 1, 42, FUZZ_SEEDS - 1):
            doc1, desc1 = self.mutate(seed)
            doc2, desc2 = self.mutate(seed)
            self.assertEqual(doc1, doc2)
            self.assertEqual(desc1, desc2)


if __name__ == '__main__':
    unittest.main()
