FIVELOG10TWO = 1.5051499783199059760686944736225

class mp3Info:
    def __init__(self, *args, **kwds):
        self.trackVolume = -1
        self.albumVolume = -1
        self.maxAmp = -1
        self.maxGain = -1
        self.minGain = -1

    def xDBGain(self, volume, target):
        if volume == -1: return -1
        return target - volume

    def xMP3Gain(self, volume, target):
        if volume == -1: return -1
        return round((target - volume) / FIVELOG10TWO)

    def xMP3DBGain(self, volume, target):
        if volume == -1: return -1
        return xMP3Gain(volume,target) * FIVELOG10TWO

    def clipping(self):
        return self.maxAmp > 1.0

    def willClip(self, volume, target):
        if volume == -1: return False
        return self.maxAmp * pow(2.0, (xMP3Gain(volume, target) / 4.0)) > 1.0

    def trackDBGain(self, target):
        return self.xDBGain(self.trackVolume, target)

    def trackMP3Gain(self, target):
        return self.xMP3Gain(self.trackVolume, target)

    def trackMP3DBGain(self, target):
        return self.xMP3DBGain(self.trackVolume, target)

    def albumDBGain(self, target):
        return self.xDBGain(self.albumVolume, target)

    def albumMP3Gain(self, target):
        return self.xMP3Gain(self.albumVolume, target)

    def albumMP3DBGain(self, target):
        return self.xMP3DBGain(self.albumVolume, target)

    def trackClip(self, target):
        return self.willClip(self.trackVolume, target)

    def albumClip(self, target):
        return self.willClip(self.albumVolume, target)
