<?php

$pageTitle = "MP3Gain";

include("start.php");
//include("quotes.php");


?>
<p>
<strong>Tired of reaching for your volume knob every time your mp3 player changes to a new song?</strong><br />
MP3Gain analyzes and adjusts mp3 files so that they have the same volume.</p>
<p>
MP3Gain does <i>not</i> just do 
<a HREF="http://replaygain.hydrogenaudio.org/faq_norm.html">peak normalization</a>, 
as many normalizers do. Instead, it does some <a HREF="http://replaygain.hydrogenaudio.org">statistical
analysis</a> to determine how loud the file actually <i>sounds</i> to the human ear.<br />
Also, the changes MP3Gain makes are completely lossless.
There is no quality lost in the change because the program adjusts the mp3 file directly, 
without decoding and re-encoding.</p>
<hr />
<p>
<strong>Latest news:</strong><br />
<em>25 May 2009</em><br />
The people who made the program SuperMp3Normalizer have chosen to re-name their product "MP3Gain PRO". I had nothing to do with this product, so don't email me support questions ;)
</p>
<hr />
<p>
<em>10 May 2009</em><br />
Zan Smogavc and his friend have translated MP3Gain into <a href="lang/Slovenscina.mp3gain.ini">Slovenian</a>.
</p>
<hr />
<p>
<em>21 Apr 2009</em><br />
Pierre le Lidgeu has updated both the <a href="help/mp3gain-french.zip">French Help file</a> and the <a href="lang/French.mp3gain.ini">French translation file</a> for version 1.2.5.
</p>
<hr />
<p>
<em>5 Feb 2009</em><br />
"REIKA" has translated the <a href="help/mp3gain-japanese.zip">Help file into Japanese</a>.
</p>
<hr />
<p>
<em>9 Jan 2009</em><br />
Luiz Gaspar has updated the <a href="lang/Portugues_Brasil.mp3gain.ini">Brazilian Portuguese</a> translation.
</p>
<hr />
<p>
<em>28 Mar 2008</em><br />
Thomas Dieffenbach has created a <a href="http://sourceforge.net/projects/easymp3gain">Linux GUI</a> for MP3Gain. It just went beta, so check it out and give him feedback
</p>
<hr />
<p>
<em>25 Dec 2007</em><br />
Wow, people are still <a href="translation.php">translating</a> MP3Gain!<br />
Just added <a href="lang/Thai.mp3gain.ini">Thai</a>.
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
<hr />
<p>
<em>09 January 2005</em><br />
Well, that was a quick bug discovery ;)<br />
AACGain 1.1 <strong>does</strong> work with the latest MP3GainGUI, but it incorrectly reports an error even after a
successful run. Dave is releasing version 1.2 very soon.<br />
Also, Dave and I will hopefully be merging the code in the near future, so AAC support will be completely integrated into
MP3Gain. We'll keep you posted.
</p>
<hr />
<p>
<em>08 January 2005</em><br />
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
<hr />
<p>
<em>16 November 2004</em><br />
<strong>Java GUI</strong>: Samuel Audet has whipped up a simple <a href="http://step.polymtl.ca/~guardia/javamp3gain.php">java GUI for mp3gain</a>. So for you non-Windows users who want a GUI but can't wait for my initial wxWidgets version, you now have another option. As a reminder, Mac users also still have <a href="http://homepage.mac.com/beryrinaldo/AudioTron/MacMP3Gain/">MacMP3Gain</a>, upon which this new JavaMP3Gain was based.<br />
</p>
<hr />
<p>
<em>12 November 2004</em><br />
Added some new <a href="translation.php">translations</a>: <a href="Srpski.mp3gain.ini">Serbian</a> and an updated <a href="Bulgarian.mp3gain.ini">Bulgarian</a>.<br />
Also added a new <a href="help/mp3gain-bulgarian.zip">Bulgarian Help file</a>.
</p>
<hr />
<p>
<em>02 November 2004</em><br />
Okay, the workaround mentioned below is out. Version 1.3.2 Beta has it.<br />
Also, I stuck some of the non-unicode improvements into the Stable version. So now Version 1.2.4 is the recommended version for most users.<br />
Also also, there was a bug in the DOS 1.4.4 code. It's fixed. Grab version 1.4.5.
</p>
<hr />
<p>
<em>03 October 2004</em><br />
Argh. I fixed the lower-case naming thing, but apparently in some cases the latest beta version is shortening the file names.<br />
So do <b>NOT</b> use the beta version unless you either
<ol>
<li>really need Unicode support, </li>
<li>want to experiment and help me figure out under exactly what circumstances the file names are shortened,</li>
<li>or you're feeling lucky</li>
</ol>
I have a kludgy workaround, but I'm trying to figure out the exact cause of the problem in the first place. Either way, I'll have version 1.3.2 out in a little while.
</p>
<hr />
<p>
<em>13 September 2004</em><br />
New 1.3.1 Beta. Someone noticed an annoying bug in 1.3.0: File names were getting reset to lower-case after running MP3Gain on them.<br />
For example, "HiThere.mp3" would become "hithere.mp3".<br/>
That bug has been fixed in 1.3.1.
</p>
<hr />
<p>
<em>07 September 2004:</em><br />
Version 1.2.3 is now officially a "stable" version. Version 1.3.0 is a new "beta" version.<br />
New features in 1.3.0:
<ul>
	<li><strong>EXTREMELY</strong> limited Unicode support-- basically just enough to get by. Unicode characters in a file name
		will show up as "?"</li>
	<li>Double-clicking on an mp3 in the list will open it in your default mp3 player. (Right-clicking and selecting "Play" works, too)</li>
</ul>
That's pretty much it.
</p>
<p>But my frustration with Visual Basic (which is what I wrote the GUI in) has finally reached critical mass.
Visual Basic does not like Unicode. Well, it doesn't like <em>displaying</em> Unicode.<br />
So I've decided to start over from scratch. The really cool part is that I'm using wxWidgets, which means I can write the code
once and compile the <strong>GUI</strong> for Windows, Linux, and Mac. (Mac users, keep in mind that a <a href="http://homepage.mac.com/beryrinaldo/AudioTron/MacMP3Gain/">MacMP3Gain</a> already exists)<br />
</p>
<p>I will also be integrating the back end and the GUI code into a single file. Don't worry, you'll still be able to use the 
command-line options if you want to, although I'll probably modify the actual parameters themselves so that they make more sense.
</p>
<p>
Oh, and I did make one tiny addition to the command-line version of mp3gain, which is now version 1.4.4:<br />
If you specify the "-r" parameter ("apply track gain"), then mp3gain skips all "Album" processing. In previous versions,
if you had multiple mp3 files specified in the command line, then mp3gain assumed you wanted to do Album processing
on all of the files in the list.<br />
Thanks to Len Trigg for pointing out how this newer method makes more sense, and even suggesting the exact code changes.
</p>
<?php 
    include("end.php");
?>
