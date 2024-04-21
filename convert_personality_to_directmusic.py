#!/usr/bin/python3
# WARNING: VERY INCOMPLETE, DOES NOT WORK YET!!!
#
# Converts a pre-DirectMusic personality to a DirectMusic chordmap.
# DirectMusic Producer can open the style (STY) files used by pre-DirectMusic
# software, but DirectMusic Producer cannot open personality and chordmap files.
# (Identify the versions used.)

import self_documenting_struct as struct

TEXT_ENCODING = 'latin-1'
def read_fixed_length_string(stream, length):
    return stream.read(length).rstrip(b'\x00').decode('latin-1')

## An "outer" chunk that has a FourCC.
class Chunk:
    ## Initializes the chunk.
    ## \param[in] stream - A binary stream positioned at the start of the chunk.
    def __init__(self, stream):
        self.fourcc = stream.read(4)
        self.length = struct.unpack.uint32_be(stream)
        self.start = stream.tell()

    ## \return The position in the file where the chunk ends.
    @property
    def end(self):
        return self.start + self.length

## An "inner" chunk, which always occurs within an "outer" chunk
## that does not have a FourCC.
class Subchunk:
    ## Initializes the subchunk. Subchunks don't have a FourCC.
    ## \param[in] stream - A binary stream positioned at the start of the chunk.
    def __init__(self, stream):
        self.length = struct.unpack.uint32_be(stream)
        self.start = stream.tell()

    ## \return The position in the file where the chunk ends.
    @property
    def end(self):
        return self.start + self.length    

## Always the first chunk in files I have observed.
class REPs(Chunk):
    def __init__(self, stream):
        super().__init__(stream)
        self.info = PersonalityInfo(stream)

        self.necs = []
        self.pnss = []
        reading_pnss = False
        while stream.tell() < self.end:
            if not reading_pnss:
                start = stream.tell()
                try:
                    self.necs.append(NECs(stream))
                except AssertionError:
                    stream.seek(start)
                    reading_pnss = True
            else:
                self.pnss.append(PNSs(stream))

class PersonalityInfo(Subchunk):
    def __init__(self, stream):
        # READ THE METADATA.
        super().__init__(stream)
        self.guid = stream.read(0x10)
        # Scale associated with the chordmap.
        # Each of the lower 24 bits represents a semitone,
        # starting with the root at the least significant bit,
        # and the bit is set if the note is in the scale.
        self.scale_pattern = stream.read(0x04)
        # Name of the chordmap, used in the object description when the chordmap is loaded.
        self.name = stream.read(0x14).rstrip(b'\x00').decode('latin-1')
        # TODO: I don't see the description of these in the spec.
        self.long_description = stream.read(0x50).rstrip(b'\x00').decode('latin-1')
        self.short_description = stream.read(0x14).rstrip(b'\x00').decode('latin-1')

        # READ THE SUBCHORD DATABASE.
        index = 0
        self.chord_definitions = []
        while stream.tell() < self.end:
            chord_definition = DMUS_IO_CHORDMAP_SUBCHORD(stream)
            if chord_definition.finished:
                break
            else:
                self.chord_definitions.append(chord_definition)
                index += 1

        # READ MORE METADATA.
        self.unk2 = stream.read(0x04)
        self.short_description_2 = stream.read(0x14).rstrip(b'\x00').decode('latin-1')
        self.long_description_2 = stream.read(0x14).rstrip(b'\x00').decode('latin-1')
        self.filename = stream.read(0x08).rstrip(b'\x00').decode('latin-1')
        self.unk3 = stream.read(0x06)

class DMUS_IO_CHORDMAP_SUBCHORD:
    def __init__(self, stream, prelude = 8):
        ## This is all zeros in all the examples I have seen.
        self.zeros = stream.read(prelude)
        if self.zeros != (b'\x00' * prelude):
            self.finished = True
            return
        else:
            self.finished = False

        ## Notes in the chord (same as in DirectMusic).
        ## Each of the lower 24 bits represents a semitone,
        ## starting with the root at the least significant bit,
        ## and the bit is set if the note is in the chord.
        self.chord_pattern = stream.read(4)

        ## A string that describes the structure of this chord.
        ## For example, "+M7" or "o7M7".
        ## Merely a descriptive, human-readable string and has 
        ## no relation to the actual notes defined for the chord.
        self.description = stream.read(0x0c).rstrip(b'\xcd').rstrip(b'\x00').decode('latin-1')

        ## The index of this chord in the personality file.
        self.index = struct.unpack.uint16_be(stream)

        ## Root of the chord (same as in DirectMusic),
        ## where 0 is the lowest C in the range and 
        ## 23 is the top B.
        self.chord_root = struct.unpack.uint8(stream)

        ## Root of the scale (same as in DirectMusic),
        ## where 0 is the lowest C in the range and 
        ## 23 is the top B.
        self.scale_root = struct.unpack.uint8(stream)

        ## Reserved for future use.
        self.flags = stream.read(2)

        ## Bit field showing which levels are supported by this subchord.
        ## Each part in a style is assigned a level, and this chord is used
        ##  only for parts whose levels are contained in this member.
        self.levels = struct.unpack.uint32_be(stream)

        ## Notes in the scale (same as in DirectMusic).
        ## Each of the lower 24 bits represents a semitone,
        ## starting with the root at the least significant bit,
        ## and the bit is set if the note is in the scale.
        self.scale_pattern = stream.read(4)

        ## Points in the scale at which inversions can occur.
        ## Bits that are off signify that the notes in the interval cannot be inverted.
        self.inversion_pattern = stream.read(4)

class NECs(Chunk):
    def __init__(self, stream):
        super().__init__(stream)
        assert self.fourcc == b'NECs'
        self.chord = NECsChordSubchunk(stream)
        self.index = struct.unpack.uint16_le(stream)
        self.lxns = LXNs(stream)

class NECsChordSubchunk(Subchunk):
    def __init__(self, stream):
        super().__init__(stream)
        self.chord = DMUS_IO_CHORDMAP_SUBCHORD(stream, prelude = 4)
        self.flags = struct.unpack.uint32_le(stream)

class LXNs(Chunk):
    def __init__(self, stream):
        super().__init__(stream)
        assert self.fourcc == b'LXNs'
        self.record_length = struct.unpack.uint32_be(stream)
        self.subchunks = []
        while stream.tell() < self.end:
            self.subchunks.append(stream.read(self.record_length))

class PNSs(Chunk):
    def __init__(self, stream):
        print(stream.tell())
        super().__init__(stream)
        assert self.fourcc == b'PNSs'
        self.unk1 = stream.read(0x4)
        self.chords = []
        for _ in range(3):
            self.chords.append(DMUS_IO_CHORDMAP_SUBCHORD(stream))
        self.data = stream.read(0x0c)

with open('.binaries/Interactive Music/MIXO.PER', 'rb') as f:
    r = REPs(f)
    print("done")