#include <math.h>
#include "mp3Info.h"

static int round_dbl(double X)
{
    int Floor, Ceil;
    
    Floor=(int)(floor(X)); Ceil=(int)(ceil(X));
    if (X-Floor<Ceil-X) return(Floor);
    else return(Ceil);
}

float mp3Info::xDBGain(float volume, float target) {
	return volume == -1 ? -1 : (target - volume);
}

int mp3Info::xMP3Gain(float volume, float target) {
	return volume == -1 ? -1 : round_dbl((target - volume) / FIVELOG10TWO);
}

float mp3Info::xMP3DBGain(float volume, float target) {
	return volume == -1 ? -1 : ((float)(xMP3Gain(volume, target)) * FIVELOG10TWO);
}

bool mp3Info::clipping() {
	return (maxAmp > 1.0);
}

bool mp3Info::willClip(float volume, float target) {
	return volume == -1 ? false : (maxAmp * pow(2.0,((double)xMP3Gain(volume, target) / (double)4.0)) > 1.0);
}
