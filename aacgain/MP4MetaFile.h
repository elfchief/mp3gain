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

//MP4MetaFile extends MP4V2 class MP4File as follows:
//
//1) Delete a free-form metadata tag.
//
//2) Modify any 8 bits of a sample
//
//3) Get file size, which may grow when Modify() is called
//
//4) Exposes the protected member function MP4File::TempFileName
//
//5) Preserves original property values for bufferSizeDB, maxBitrate and avgBitrate.
//
//6) Preserves 'free' atom between 'moov' and 'mtda' atoms in files created by iTunes
#ifndef __MP4_META_FILE_H__
#define __MP4_META_FILE_H__

#pragma warning( push )
#pragma warning( disable : 4100 4244 )
#include "mp4common.h"
#pragma warning( pop )

class MP4MetaFile : public MP4File
{
public:
    MP4MetaFile(u_int32_t verbosity = 0);

    bool DeleteMetadataFreeForm(char *pName);
    void ModifySampleByte(MP4TrackId trackId, MP4SampleId sampleId, u_int8_t byte,
                          u_int32_t byteOffset, u_int8_t bitOffset);
    u_int64_t GetFileSize();
    const char* TempFileName();
    u_int64_t GetFreeAtomSize();

    //overrides of MP4File member functions
    void Close();
    void FinishWrite();
    void Optimize(const char* orgFileName, const char* newFileName=NULL, u_int32_t freeAtomSize=0);
};

#endif //__MP4_META_FILE_H__