#ifndef MP3INFO_H
#define MP3INFO_H

#include <wx/wx.h>

#define FIVELOG10TWO 1.5051499783199059760686944736225

class mp3Info {
private:
	bool willClip(float volume, float target);
	float xDBGain(float volume, float target);
	int xMP3Gain(float volume, float target);
	float xMP3DBGain(float volume, float target);

public:
	wxString path;
	wxString filename;
	float trackVolume;
	float albumVolume;
	double maxAmp;
	int maxGain;
	int minGain;

	mp3Info() {
		trackVolume = -1;
		albumVolume = -1;
		maxAmp = -1;
		maxGain = -1;
		minGain = -1;
	};

	mp3Info(const wxString newPath) {
		path = newPath;
		trackVolume = -1;
		albumVolume = -1;
		maxAmp = -1;
		maxGain = -1;
		minGain = -1;
	};

	~mp3Info() {
		//wxMessageBox(wxString::Format(wxT("Deleting object %s"),path.c_str()));
	};

	float trackDBGain(float target) { return xDBGain(trackVolume, target); };
	int trackMP3Gain(float target) { return xMP3Gain(trackVolume, target); };
	float trackMP3DBGain(float target) { return xMP3DBGain(trackVolume, target); };
	float albumDBGain(float target) { return xDBGain(albumVolume, target); };
	int albumMP3Gain(float target) { return xMP3Gain(albumVolume, target); };
	float albumMP3DBGain(float target) { return xMP3DBGain(albumVolume, target); };
	bool clipping();
	bool trackClip(float target) { return willClip(trackVolume, target); };
	bool albumClip(float target) { return willClip(albumVolume, target); };

};


#endif
