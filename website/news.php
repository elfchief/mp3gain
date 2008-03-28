<?php

$pageTitle = "MP3Gain News";

include("start.php");

?>
<h1 class="hide">News</h1>
<p>
<em>28 Mar 2008</em><br />
Thomas Dieffenbach has created a <a href="http://sourceforge.net/projects/easymp3gain">Linux GUI</a> for MP3Gain. It just went beta, so check it out and give him feedback
</p>
<hr />
<p>
<em>25 Dec 2007</em><br />
Wow, people are still <a href="translation.php">translating</a> MP3Gain!<br />
Just added <a href="help/Thai.mp3gain.ini">Thai</a>.
</p>
<hr />
<p>
<em>19 March 2005</em><br />
Just a reminder that the new AAC part of mp3gain is <strong>experimental</strong>. It's simply
newer, so problems are still being found (and fixed). Use it at your own risk, and I'd suggest
backing up your files first.
</p>
<hr />
<p>
<em>10 January 2005</em><br />
Bug fixed. If you use AACGain with the MP3Gain GUI, make sure you get
<a href="http://www.rarewares.org/aac.html">aacgain version 1.2</a> or later.</p>
</p>
<p>
<em>09 January 2005</em><br />
Well, that was a quick bug discovery ;)<br />
AACGain 1.1 <strong>does</strong> work with the latest MP3GainGUI, but it incorrectly reports an error even after a
successful run. Dave is releasing version 1.2 very soon.<br />
Also, Dave and I will hopefully be merging the code in the near future, so AAC support will be completely integrated into
MP3Gain. We'll keep you posted.
</p>
<p><em>08 January 2005</em><br />
<strong>AACGain</strong>: Dave Lasker has added AAC support to mp3gain.exe. He wrote aacgain.exe specifically so it would
work with the existing MP3GainGUI without too much trouble.<br />
To get it all to work, go <a href="download.php">download the latest MP3Gain</a> (either "1.2.5 Stable" or "1.3.4 Beta").
Then <a href="http://www.rarewares.org/aac.html">download AACGain</a>. Un-zip aacgain.exe, re-name it to "mp3gain.exe",
and move it into the MP3Gain folder, copying over the existing mp3gain.exe.<br />
That's all you have to do. Now MP3Gain should handle AAC files (.m4a or .mp4).</p>
<p>Please note that aacgain will not work on DRM-encoded files (i.e. music you buy from the iTunes store).
It should work just fine with AAC file you create yourself using iTunes, though.
</p>
<p>And a technical note for command-line users: As part of coordinating this release with Dave, I've finally fixed
the program return codes in mp3gain.exe to match what everyone else in the world does. So as of version 1.4.6,
0 means success, and non-zero means failure.
</p>
<p><em>16 November 2004</em><br />
<strong>Java GUI</strong>: Samuel Audet has whipped up a simple <a href="http://step.polymtl.ca/~guardia/javamp3gain.php">java GUI for mp3gain</a>. So for you non-Windows users who want a GUI but can't wait for my initial wxWidgets version, you now have another option. As a reminder, Mac users also still have <a href="http://homepage.mac.com/beryrinaldo/AudioTron/MacMP3Gain/">MacMP3Gain</a>, upon which this new JavaMP3Gain was based.<br />
</p>
<p><em>12 November 2004</em><br />
Added some new <a href="translation.php">translations</a>: <a href="Srpski.mp3gain.ini">Serbian</a> and an updated <a href="Bulgarian.mp3gain.ini">Bulgarian</a>.<br />
Also added a new <a href="help/mp3gain-bulgarian.zip">Bulgarian Help file</a>.
</p>
<p><em>02 November 2004</em><br />
Version 1.3.2 (Beta) is out. I still only recommend it if you really need Unicode support.<br />
Version 1.2.4 (stable) now includes the extra little tweaks that 1.3.1 had, but without the Unicode stuff. For instance, double-clicking on a file in the list opens the mp3 file in your default player.
</p>
<p><em>03 October 2004</em><br />
Sigh. Version 1.3.1 is still buggy. Hence the "beta" name. Don't use it unless you really need Unicode support.
</p>
<p><em>13 September 2004</em><br />
Version 1.3.1 is the new beta. It's exactly the same as 1.3.0, plus a bug fix:<br/>
In 1.3.0, file names were sometimes re-named to lower case after running MP3Gain on them. For example "HiThere.mp3" would become "hithere.mp3"
</p>
<p><em>07 September 2004</em><br />
Version 1.2.3 is officially "stable".<br />
Version 1.3.0 is the new beta version including limited Unicode file name support
</p>
<p>Version 1.2.3 is now available on the Downloads page. No real changes since version 1.2.2 (just a tiny, almost inconsequential bug fix). But the Help file is now up-to-date, and I <strong>will</strong> be making more changes very soon.</p>

<? include("end.php"); ?>
