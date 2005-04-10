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

//Portions of this file are derived from faad2 file frontend/main.c

//Thanks to Prakash Punoor for help making it portable

#include "neaacdec.h"
#include "aacgain.h"
#include "aacgaini.h"
//following header #includes mpeg4ip/include/mpeg4ip.h which #includes common c-lib include files
#include "MP4MetaFile.h"

#ifdef WIN32
#include <sys/utime.h>
#else
#include <utime.h>
#endif

#ifndef max
#define max(X,Y) ((X)>(Y)?(X):(Y))
#endif

#ifndef min
#define min(X,Y) ((X)<(Y)?(X):(Y))
#endif

#define SQRTHALF            0.70710678118654752440084436210485

#define NEXT_SAMPLE(dest)\
{\
    decode_t s;\
    s = *sample++;\
    dest = s * 32768.0;\
    if (s < 0) s = -s;\
    theGainData->peak = max(s, theGainData->peak);\
}

GainDataPtr theGainData;
static const MP4TrackId theTrack = 1;
static const u_int32_t verbosity = MP4_DETAILS_ERROR|MP4_DETAILS_WARNING;

//replay_gain tags
static char *RGTags[num_rg_tags] = 
{
    "replaygain_track_gain",
    "replaygain_album_gain",
    "replaygain_track_peak",
    "replaygain_album_peak",
    "replaygain_track_minmax",
    "replaygain_album_minmax",
    "replaygain_undo"
};

typedef struct
{
    struct stat savedAttributes;
} PreserveTimestamp, *PreserveTimestampPtr;

void modifyGain(MP4MetaFile* mp4MetaFile, int delta, GainFixupPtr gf)
{
    uint8_t new_gain = gf->orig_gain + (uint8_t)delta;

    //update global_gain
    mp4MetaFile->ModifySampleByte(theTrack, gf->sampleId, new_gain, 
        gf->sample_offset, gf->bit_offset);
}

static int AACAnalyze(void *sample_buffer, long numSamples, unsigned char channels, 
                      int compute_gain)
{
    decode_t *samples = (decode_t *)sample_buffer;
    decode_t *sample = samples;
    rg_t *left_samples;
    rg_t *right_samples;
    long i;
    
    left_samples = new rg_t[numSamples];
    if (!left_samples)
        return 1;
    if ((channels == 2) || (channels == 6))
    {
        right_samples = new rg_t[numSamples];
    } else {
        right_samples = NULL;
    }

    switch (channels)
    {
    case 1:
        for (i=0; i<numSamples; i++)
        {
            NEXT_SAMPLE(left_samples[i])
        }
        break;
    case 2:
        for (i=0; i<numSamples; i++)
        {
            NEXT_SAMPLE(left_samples[i])
            NEXT_SAMPLE(right_samples[i])
        }
       break;
    case 6:
        for (i=0; i<numSamples; i++)
        {
            //faad2 gives samples in following order: c,l,r,bl,br,lfe
            decode_t c, l, r, bl, br, lfe;
            NEXT_SAMPLE(c);
            NEXT_SAMPLE(l);
            NEXT_SAMPLE(r);
            NEXT_SAMPLE(bl);
            NEXT_SAMPLE(br);
            NEXT_SAMPLE(lfe);
            left_samples[i] = l + c*SQRTHALF + bl*SQRTHALF + lfe;
            right_samples[i] = r + c*SQRTHALF + br*SQRTHALF + lfe;
        }
        break;
    default:
        for (i=0; i<numSamples; i++)
        {
            int j;
            decode_t sum = 0;
            for (j=0; j<channels; j++)
            {
                decode_t samp;
                NEXT_SAMPLE(samp)
                sum += samp;
            }
            left_samples[i] = (rg_t)sum / (rg_t)channels;
        }
    }

    if (compute_gain)
        AnalyzeSamples(left_samples, right_samples, numSamples, right_samples ? 2 : 1);

    delete [] left_samples;
    if (right_samples)
        delete [] right_samples;

    return 0;
}

static int parseMp4File(GainDataPtr gd, ProgressCallback reportProgress, int compute_gain)
{
    NeAACDecHandle hDecoder = gd->hDecoder;
    unsigned char *buffer;
    unsigned int buffer_size;
    void *sample_buffer;
    NeAACDecFrameInfo frameInfo;
    MP4MetaFile* mp4MetaFile = (MP4MetaFile*)gd->mp4MetaFile;

    unsigned long sampleId, numSamples;
    int percent, old_percent = -1;

    theGainData = gd;
    gd->GainHead = gd->GainTail = NULL;
    frameInfo.error = 0;
    numSamples = mp4MetaFile->GetTrackNumberOfSamples(theTrack);

    for (sampleId = 1; sampleId <= numSamples; sampleId++)
    {
        /* get acces unit from MP4 file */
        buffer = NULL;
        buffer_size = 0;

        //set sampleId for use by syntax.c
        gd->sampleId = sampleId;

        try 
        {
            mp4MetaFile->ReadSample(theTrack, sampleId, (u_int8_t**)(&buffer), (u_int32_t*)(&buffer_size));
        } catch (MP4Error* e)
        {
            e->Print();
            fprintf(stderr, "Reading from MP4 file failed. \n");
            NeAACDecClose(hDecoder);
            free (e);
            return 1;
        }

        sample_buffer = aacgainDecode(hDecoder, &frameInfo, buffer, buffer_size);
        if (gd->analyze && (frameInfo.error == 0) && (frameInfo.samples > 0))

        {
            AACAnalyze(sample_buffer, frameInfo.samples/gd->channels, gd->channels, compute_gain);
        }

        if (buffer) free(buffer);

        percent = min((int)(sampleId*100)/numSamples, 100);
        if (reportProgress && (percent > old_percent))
        {
            old_percent = percent;
            reportProgress(percent, (unsigned int)mp4MetaFile->GetFileSize());
        }

        if (frameInfo.error > 0)
        {
            fprintf(stderr, "Error: invalid file format %s, code=%d\n",
                gd->mp4file_name, frameInfo.error);
            gd->abort = 1;
            break;
        }
    }

    return frameInfo.error;
}

static MP4MetaFile *PrepareToWrite(GainDataPtr gd)
{
    MP4MetaFile* mp4MetaFile = (MP4MetaFile*)gd->mp4MetaFile;

    if (!gd->open_for_write)
    {
		if (gd->use_temp)
		{
			//if we are using a temp file, create it now...
			gd->temp_name = mp4MetaFile->TempFileName(gd->mp4file_name);
			FILE *tmpFile = fopen(gd->temp_name, "wb");
			if (!tmpFile)
			{
				fprintf(stderr, "Error: unable to create temporary file %s\n", gd->temp_name);
				exit(1);
			}
			//close the MP4MetaFile and reopen as stdio file
			mp4MetaFile->Close();
			delete mp4MetaFile;
			FILE *inFile = fopen(gd->mp4file_name, "rb");
			if (!inFile)
			{
				fprintf(stderr, "Error: unable to reopen file %s to create temporary file\n",
					gd->mp4file_name);
				exit(1);
			}

			//copy the original file to the temp file
			static const u_int32_t blockSize = 4096;
			u_int8_t *buffer = new u_int8_t[blockSize];
			for (;;)
			{
				int bytesRead = fread(buffer, 1, blockSize, inFile);
				if (bytesRead)
					fwrite(buffer, 1, bytesRead, tmpFile);
				if (bytesRead < blockSize)
					break;
			}
			fclose(inFile);
			fclose(tmpFile);
			delete buffer;
			try
			{
				gd->mp4MetaFile = mp4MetaFile = new MP4MetaFile(verbosity);
				mp4MetaFile->Modify(gd->temp_name);
			} catch(MP4Error *e) {
				fprintf(stderr, "Unable to open file %s for writing.\n", gd->temp_name);
				free(e);
				exit(1);
			}
		} else {
			//otherwise open the file for writing
			try
			{
				mp4MetaFile->Close();
				delete mp4MetaFile;
				gd->mp4MetaFile = mp4MetaFile = new MP4MetaFile(verbosity);
				mp4MetaFile->Modify(gd->mp4file_name);
			} catch(MP4Error *e) {
				fprintf(stderr, "Unable to open file %s for writing. It may be in use\n"
					"by another program\n", gd->mp4file_name);
				free(e);
				exit(1);
			}
		}
        gd->open_for_write = 1;
    }

    return mp4MetaFile;
}

int aac_open(char *mp4_file_name, int use_temp, int preserve_timestamp, AACGainHandle *gh)
{
    FILE* mp4_file;
    GainDataPtr gd;
    unsigned char header[8];
    size_t file_name_len;
    MP4MetaFile* mp4MetaFile;
    PreserveTimestampPtr pt = NULL;

    *gh = NULL;

    //In order to allow processed files to play on iPod Shuffle, which is extremley sensitive to
    // file format, we always use a temp file. This runs the MP4File::Optimize function,
    // which rewrites the processed file in the connanical order.
    use_temp = true;

    file_name_len = strlen(mp4_file_name);
    if ((file_name_len >= 5) && (strcmp(mp4_file_name + file_name_len - 4, ".m4p") == 0))
    {
        fprintf(stderr, "Error: DRM protected file %s is not supported.\n", mp4_file_name);
        return 1;
    }

    mp4_file = fopen(mp4_file_name, "rb");
    if (!mp4_file)
    {
        //caller's responsibility to give error message so we can use aac_open to test for aac file
        return 0;
    }

    fread(header, 1, 8, mp4_file);
    fclose(mp4_file);
    if (header[4] != 'f' || header[5] != 't' || header[6] != 'y' || header[7] != 'p')
    {
        //no error - use this to tell if a file is mp3 or mp4
        return 0;
    }

    if (preserve_timestamp)
    {
        pt = new PreserveTimestamp;
        stat(mp4_file_name, &pt->savedAttributes);
    }

    gd = new GainData;
    gd->mp4MetaFile = NULL;
    gd->analyze = 0;
    gd->use_temp = use_temp;
    gd->open_for_write = 0;
    gd->gain_read = 0;
    gd->peak = 0;
    gd->hDecoder = NULL;
    gd->abort = 0;
    gd->preserve_timestamp = pt;
    gd->GainHead = NULL;

    gd->mp4file_name = strdup(mp4_file_name);
    gd->temp_name = NULL;

    try
    {
        mp4MetaFile = new MP4MetaFile(verbosity);
        mp4MetaFile->Read(mp4_file_name);
        if (mp4MetaFile->GetNumberOfTracks(MP4_AUDIO_TRACK_TYPE) != 1)
        {
            fprintf(stderr, "File must contain a single audio track.\n");
            throw new MP4Error();
        }
        gd->free_atom_size = (u_int32_t)mp4MetaFile->GetFreeAtomSize();
        gd->mp4MetaFile = mp4MetaFile;
    } catch (MP4Error* e)
    {
        /* unable to open file */
        fprintf(stderr, "Error opening file: %s\n", gd->mp4file_name);
        gd->abort = 1;
        aac_close(gd);
        free (e);
        return 1;
    }

    NeAACDecHandle hDecoder;
    NeAACDecConfigurationPtr config;
    mp4AudioSpecificConfig mp4ASC;
    unsigned char *buffer;
    unsigned int buffer_size;

    hDecoder = gd->hDecoder = NeAACDecOpen();

    /* Set configuration */
    config = NeAACDecGetCurrentConfiguration(hDecoder);
    config->outputFormat = FAAD_FMT_DOUBLE;
    config->downMatrix = 0;
    NeAACDecSetConfiguration(hDecoder, config);

    buffer = NULL;
    buffer_size = 0;
    mp4MetaFile->GetTrackESConfiguration(theTrack, (u_int8_t**)(&buffer), (u_int32_t*)(&buffer_size));
    if (NeAACDecInit2(hDecoder, buffer, buffer_size,
                    &gd->samplerate, &gd->channels) < 0)
    {
        /* If some error initializing occured, skip the file */
        if (buffer)
            free(buffer);
        fprintf(stderr, "Error: file format not recognized.\n");
        gd->abort = 1;
        aac_close(gd);
        return 1;
    }

    if (NeAACDecAudioSpecificConfig(buffer, buffer_size, &mp4ASC) >= 0)
    {
        if (mp4ASC.sbr_present_flag == 1)
        {
            free(buffer);
            fprintf(stderr, "Error: HE_AAC/SBR files are not supported.\n");
            gd->abort = 1;
            aac_close(gd);
            return 1;
        }
    }
    free(buffer);

    *gh = gd;
    return 0;
}

unsigned int aac_get_sample_rate(AACGainHandle gh)
{
    GainDataPtr gd = (GainDataPtr)gh;

    return gd->samplerate;
}

int aac_compute_gain(AACGainHandle gh, rg_t *peak, unsigned char *min_gain, 
                     unsigned char *max_gain, ProgressCallback reportProgress)
{
    int rc = 0;
    GainDataPtr gd = (GainDataPtr)gh;

    if (!gd->gain_read || !gd->analyze)
    {
	    gd->analyze = 1;
        gd->peak = 0;
        gd->min_gain = 255;
        gd->max_gain = 0;
        rc = parseMp4File(gd, reportProgress, 1);
        gd->gain_read = 1;
    }
    if (peak)
    {
        *peak = gd->peak * 32768.0;
    }
    if (min_gain)
        *min_gain = gd->min_gain;
    if (max_gain)
        *max_gain = gd->max_gain;
    
    return rc;
}

int aac_compute_peak(AACGainHandle gh, rg_t *peak, unsigned char *min_gain,
                     unsigned char *max_gain, ProgressCallback reportProgress)
{
    int rc = 0;
    GainDataPtr gd = (GainDataPtr)gh;

    if (!gd->gain_read || !gd->analyze)
    {
		gd->analyze = 1;
        gd->peak = 0;
        gd->min_gain = 255;
        gd->max_gain = 0;
        rc = parseMp4File(gd, reportProgress, 0);
        gd->gain_read = 1;
    }
    if (peak)
    {
        *peak = gd->peak * 32768.0;
    }
    if (min_gain)
        *min_gain = gd->min_gain;
    if (max_gain)
        *max_gain = gd->max_gain;
    
    return rc;
}

int aac_modify_gain(AACGainHandle gh, int left, int right,
                    ProgressCallback reportProgress)
{
    GainDataPtr gd = (GainDataPtr)gh;
    MP4MetaFile* mp4MetaFile = (MP4MetaFile*)gd->mp4MetaFile;
    GainFixupPtr gf;
    int rc = 0;

    if ((gd->channels != 2) && (left != right))
    {
        fprintf(stderr, "Error: individual channel adjustments are only supported on\n"
            "2-channel (stereo) files.\n");
        gd->abort = 1;
        return -1;
    }

    if (!gd->gain_read)
    {
        gd->analyze = 0;
        rc = parseMp4File(gd, reportProgress, 0);
        gd->gain_read = 1;
        if (rc)
        {
            gd->abort = 1;
            return rc;
        }
    }

    //test for wrap before modifying file
    gf = gd->GainHead;
    while (gf)
    {
        if (((gf->channel == 0) && 
            (((gf->orig_gain + left) < 0) || ((gf->orig_gain + left) > 255))) ||
            ((gf->channel == 1) && 
            (((gf->orig_gain + right) < 0) || ((gf->orig_gain + right) > 255))))
        {
            fprintf(stderr, "Error: Wrap while modifying gain.\n");
            gd->abort = 1;
            return -1;
        }
        gf = gf->next;
    }

    mp4MetaFile = PrepareToWrite(gd);

    gf = gd->GainHead;
    while (gf)
    {
        GainFixupPtr prev;

		//update global_gain
        modifyGain(mp4MetaFile, (gf->channel == 0) ? left : right, gf);
        prev = gf;
        gf = gf->next;
        free(prev);
    }
    gd->GainHead = NULL;

    return rc;
}

int aac_set_tag_float(AACGainHandle gh, rg_tag_e tag, rg_t value)
{
    GainDataPtr gd = (GainDataPtr)gh;
    MP4MetaFile* mp4MetaFile = PrepareToWrite(gd);
    char vstr[20];

    sprintf(vstr, "%-.2f", value);
    mp4MetaFile->SetMetadataFreeForm(RGTags[tag], (u_int8_t*)vstr, strlen(vstr));

    return 0;
}

int aac_get_tag_float(AACGainHandle gh, rg_tag_e tag, rg_t *value)
{
    GainDataPtr gd = (GainDataPtr)gh;
    MP4MetaFile* mp4MetaFile = (MP4MetaFile*)gd->mp4MetaFile;
    char *vstr;
	u_int32_t vsize;

    if (mp4MetaFile->GetMetadataFreeForm(RGTags[tag], (u_int8_t**)&vstr, &vsize))
    {
		//null terminate the value
		vstr = (char*)realloc(vstr, vsize+1);
		vstr[vsize] = '\0';

		sscanf(vstr, "%lf", value);
        free(vstr);
        return 0;
    }

    return 1;
}

int aac_set_tag_int_2(AACGainHandle gh, rg_tag_e tag, int p1, int p2)
{
    GainDataPtr gd = (GainDataPtr)gh;
    MP4MetaFile* mp4MetaFile = PrepareToWrite(gd);

    char vstr[100];

    sprintf(vstr, "%d,%d", p1, p2);
    mp4MetaFile->SetMetadataFreeForm(RGTags[tag], (u_int8_t*)vstr, strlen(vstr));

    return 0;
}

int aac_get_tag_int_2(AACGainHandle gh, rg_tag_e tag, int *p1, int *p2)
{
    GainDataPtr gd = (GainDataPtr)gh;
    MP4MetaFile* mp4MetaFile = (MP4MetaFile*)gd->mp4MetaFile;
    char *vstr;
	u_int32_t vsize;

    if (mp4MetaFile->GetMetadataFreeForm(RGTags[tag], (u_int8_t**)&vstr, &vsize))
    {
		//null terminate the value
		vstr = (char*)realloc(vstr, vsize+1);
		vstr[vsize] = '\0';

		sscanf(vstr, "%d,%d", p1, p2);
        free(vstr);
        return 0;
    }

    return 1;
}

int aac_clear_rg_tags(AACGainHandle gh)
{
    GainDataPtr gd = (GainDataPtr)gh;
    MP4MetaFile* mp4MetaFile = PrepareToWrite(gd);
    uint32_t i;

    for (i=0; i<num_rg_tags; i++)
    {
        mp4MetaFile->DeleteMetadataFreeForm(RGTags[i]);
    }

    return 0;
}

int aac_close(AACGainHandle gh)
{
    GainDataPtr gd = (GainDataPtr)gh;
    MP4MetaFile* mp4MetaFile = (MP4MetaFile*)gd->mp4MetaFile;
    int rc = 0;
    PreserveTimestampPtr pt = (PreserveTimestampPtr)gd->preserve_timestamp;
    const char *tempFileName = NULL;

    //close the faad decoder if open
    if (gd->hDecoder)
    {
        NeAACDecClose(gd->hDecoder);
    }

    //delete the gain change linked list if present
    while (gd->GainHead)
    {
        GainFixupPtr next = gd->GainHead->next;
        free(gd->GainHead);
        gd->GainHead = next;
    }

    if (mp4MetaFile)
    {
        if (gd->use_temp && gd->temp_name)
            tempFileName = mp4MetaFile->TempFileName(gd->mp4file_name);

        mp4MetaFile->Close();
        delete mp4MetaFile;
    }

    if (tempFileName)
    {
        if (!gd->abort)
        {
            //use MP4File::Optimize to undo the wasted space created by MP4File::Modify
            //send optimize output to a temp file "just in case"
            MP4MetaFile f;
            f.Optimize(gd->temp_name, tempFileName, gd->free_atom_size);

            //rename the temp file back to original name
            int rc = remove(gd->mp4file_name);
            if (rc == 0)
                rc = rename(tempFileName, gd->mp4file_name);
            if (rc)
                fprintf(stderr, "Error: attempt to create file %s failed. Your output file is named %s",
                    gd->mp4file_name, tempFileName);
            free((void*)tempFileName);
        }
        remove(gd->temp_name);
        free((void*)gd->temp_name);
    }

    if (pt)
    {
        if (!gd->abort)
        {
			struct utimbuf setTime;	
			
			setTime.actime = pt->savedAttributes.st_atime;
			setTime.modtime = pt->savedAttributes.st_mtime;
			utime(gd->mp4file_name, &setTime);
        }
        delete pt;
    }

    free(gd->mp4file_name);
    delete gd;

    return rc;
}