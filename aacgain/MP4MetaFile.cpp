/*
** aacgain - modifications to mp3gain to support mp4/m4a files
** Copyright (C) David Lasker, 2004 Altos Design, Inc.
**
** This program is free software; you can redistribute it and/or modify
** it under the terms of the GNU General Public License as published by
** the Free Software Foundation; either version 2 of the License, or
** (at your option) any later version.
**
** This program is distributed in the hope that it will be useful,
** but WITHOUT ANY WARRANTY; without even the implied warranty of
** MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
** GNU General Public License for more details.
**
** You should have received a copy of the GNU General Public License
** along with this program; if not, write to the Free Software
** Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
**/

#include "MP4MetaFile.h"

//MyMP4Track is a kluge to allow us to call protected member function
// MP4Track::GetSampleFileOffset(), and to override MP4Track::FinishWrite. 
// We do this by casting a MP4Track to MyMP4Track. I'm not sure if that is
// a C++ legal downcast, but based on my userstanding of how C++ generated
// code handles the "this" pointer as an extra parameter, it should work.
class MyMP4Track : public MP4Track
{
private:
    MyMP4Track(); //can not be instantiated, only cast

public:
    u_int64_t	GetSampleFileOffset(MP4SampleId sampleId)
    {
        return MP4Track::GetSampleFileOffset(sampleId);
    }

    //Override MP4Track::FinishWrite to preserve original 
    // bufferSizeDB, maxBitrate and avgBitrate. This preserves the
    // original values of these properties, so that iTunes displays
    // then as "whole" numbers, i.e 320KB instead of 319KB.
    void FinishWrite()
    {
	    // write out any remaining samples in chunk buffer
	    WriteChunkBuffer();
    }
};

//A similar kluge to allow us to override protected member function
// MP4RootAtom::BeginOptimalWrite() to write the extra 'free'
// atom used by iTunes
class MyMP4RootAtom : public MP4RootAtom
{
private:
    MyMP4RootAtom();

public:
    void BeginOptimalWrite(u_int32_t freeAtomSize)
    {
	    WriteAtomType("ftyp", OnlyOne);
	    WriteAtomType("moov", OnlyOne);
	    WriteAtomType("udta", Many);
        //AACGain: write the extra 'free' atom used by iTunes
        if (freeAtomSize)
        {
            MP4FreeAtom* freeAtom = new MP4FreeAtom();
            freeAtom->SetFile(m_pFile);
            freeAtom->SetSize(freeAtomSize);
            freeAtom->Write();
            delete freeAtom;
        }

	    m_pChildAtoms[GetLastMdatIndex()]->BeginWrite(m_pFile->Use64Bits("mdat"));
    }
};

MP4MetaFile::MP4MetaFile(u_int32_t verbosity)
: MP4File(verbosity)
{
}

bool MP4MetaFile::DeleteMetadataFreeForm(char *name)
{
    char s[256];
    int	i =	0;

    for (;;)
    {
        MP4BytesProperty *pMetadataProperty;
        sprintf(s, "moov.udta.meta.ilst.----[%u].name",	i);
        MP4Atom	*pTagNameAtom =	m_pRootAtom->FindAtom(s);
        if (!pTagNameAtom)
            return false;
        pTagNameAtom->FindProperty("name.metadata",	(MP4Property**)&pMetadataProperty);
        if (pMetadataProperty)
        {
            u_int8_t* pV;
            u_int32_t VSize	= 0;
            pMetadataProperty->GetValue(&pV, &VSize);
            if (VSize != 0)
            {
                if ((VSize == strlen(name)) && (memcmp((char*)pV, name, VSize) == 0))
                {
                    MP4Free(pV);
                    MP4Atom *p4dashesAtom = pTagNameAtom->GetParentAtom(); //the '----' atom
                    MP4Atom *pIlstAtom = p4dashesAtom->GetParentAtom();
                    pIlstAtom->DeleteChildAtom(p4dashesAtom);
                    delete p4dashesAtom;
                    return true;
                }
                MP4Free(pV);
            }
        }
        i++;
    }
}

void MP4MetaFile::ModifySampleByte(MP4TrackId trackId, MP4SampleId sampleId, u_int8_t byte,
                                   u_int32_t byteOffset, u_int8_t bitOffset)
{
    ProtectWriteOperation("MP4MetaFile::ModifySampleByte");

    u_int64_t sampleOffset = static_cast<MyMP4Track *>(m_pTracks[FindTrackIndex(trackId)])->
            GetSampleFileOffset(sampleId);
    u_int64_t origPosition = GetPosition();

    SetPosition(sampleOffset + byteOffset);

    if (bitOffset)
    {
        //the 8 bits span 2 bytes
        u_int8_t buf[2];
        PeekBytes(buf, 2);
        buf[0] &= (0xff << bitOffset);
        buf[0] |= (byte >> (8 - bitOffset));
        buf[1] &= (0xff >> (8 - bitOffset));
        buf[1] |= (byte << bitOffset);
        WriteBytes(buf, 2);
    } else {
        //the 8 bits is byte-aligned
        WriteBytes(&byte, 1);
    }

    SetPosition(origPosition);
}

u_int64_t MP4MetaFile::GetFileSize()
{
    return m_fileSize;
}

const char* MP4MetaFile::TempFileName()
{
    //strdup the result of MP4File::TempFileName() since
    // the string needs to outlive the class instance
    return strdup(MP4File::TempFileName());
}

//return the size of the 'free' atom used for padding between 'moov' and 'mdta'
u_int64_t MP4MetaFile::GetFreeAtomSize()
{
    u_int32_t nChildren = m_pRootAtom->GetNumberOfChildAtoms();
    while (nChildren-- > 0) 
    {
        MP4Atom* child = m_pRootAtom->GetChildAtom(nChildren);
        if (!strcmp(child->GetType(), "free"))
            return child->GetSize();
    }
    return 0;
}

//override MP4File::Close to call MP4MetaFile::FinishWrite
void MP4MetaFile::Close()
{
	if (m_mode == 'w') {
		SetIntegerProperty("moov.mvhd.modificationTime", 
			MP4GetAbsTimestamp());

		FinishWrite();
	}

	fclose(m_pFile);
	m_pFile = NULL;
}

//override MP4File::FinishWrite to call MyMP4Track::FinishWrite
void MP4MetaFile::FinishWrite()
{
	// for all tracks, flush chunking buffers
	for (u_int32_t i = 0; i < m_pTracks.Size(); i++) {
		ASSERT(m_pTracks[i]);
        ((MyMP4Track*)m_pTracks[i])->MyMP4Track::FinishWrite();
	}

	// ask root atom to write
	m_pRootAtom->FinishWrite();

	// check if file shrunk, e.g. we deleted a track
	if (GetSize() < m_orgFileSize) {
		// just use a free atom to mark unused space
		// MP4Optimize() should be used to clean up this space
		MP4Atom* pFreeAtom = MP4Atom::CreateAtom("free");
		ASSERT(pFreeAtom);
		pFreeAtom->SetFile(this);
		int64_t size = m_orgFileSize - (m_fileSize + 8);
		if (size < 0) size = 0;
		pFreeAtom->SetSize(size);
		pFreeAtom->Write();
		delete pFreeAtom;
	}
}

//override MP4File::Optimize to preserve extra 'free' atom used by iTunes
void MP4MetaFile::Optimize(const char* orgFileName, const char* newFileName, u_int32_t freeAtomSize)
{
	m_fileName = MP4Stralloc(orgFileName);
	m_mode = 'r';

	// first load meta-info into memory
	Open("rb");
	ReadFromFile();

	CacheProperties();	// of moov atom

	// now switch over to writing the new file
	MP4Free(m_fileName);

	// create a temporary file if necessary
	if (newFileName == NULL) {
		m_fileName = MP4Stralloc(TempFileName());
	} else {
		m_fileName = MP4Stralloc(newFileName);
	}

	FILE* pReadFile = m_pFile;
	m_pFile = NULL;
	m_mode = 'w';

	Open("wb");

	SetIntegerProperty("moov.mvhd.modificationTime", 
		MP4GetAbsTimestamp());

	// writing meta info in the optimal order
    //AACGain: call MyMP4RootAtom::BeginOptimalWrite to write the extra 'free' atom used by iTunes
	((MyMP4RootAtom*)m_pRootAtom)->BeginOptimalWrite(freeAtomSize);

	// write data in optimal order
	RewriteMdat(pReadFile, m_pFile);

	// finish writing
	((MP4RootAtom*)m_pRootAtom)->FinishOptimalWrite();


	// cleanup
	fclose(m_pFile);
	m_pFile = NULL;
	fclose(pReadFile);

	// move temporary file into place
	if (newFileName == NULL) {
		Rename(m_fileName, orgFileName);
	}
}