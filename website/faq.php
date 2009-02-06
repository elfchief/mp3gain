<?php

$pageTitle = "MP3Gain FAQ";

include("start.php");

	$faqs = array(	array ("General questions",
							array (
									array ("start","What does MP3Gain do?",
	"MP3Gain automatically adjusts mp3s so that they all have the same volume"),
									array ("peak","You mean MP3Gain normalizes mp3 files?",
	"Yes, but MP3Gain does not use &quot;peak amplitude&quot; normalization as many &quot;normalizers&quot; do.
	Audio files with <i>very</i> different peak amplitudes can still sound to the human ear as though they're the same volume.<br />
	Instead, MP3Gain uses David Robinson's <a href=\"http://replaygain.hydrogenaudio.org\">Replay Gain</a> algorithm to calculate how
	loud the file actually sounds to a human's ears.
	<p>To hear the difference between &quot;maximizing&quot; (peak normalization) and Replay Gain volume normalization,
	<ol>
		<li>Download this <a href=\"maxampdemo.zip\">sample file</a>
		<li>Unzip the two mp3 files, noting their current maximum amplitudes as indicated in the filenames
		<li>Open MP3Gain
		<li>Go to &quot;Options -&gt; Advanced...&quot; and make sure the &quot;Enable Maximizing features&quot; option is checked
		<li>Set the &quot;Target Normal Volume&quot; to 92.0 dB
		<li>Click &quot;Add Files,&quot; and add the two unzipped mp3 files
		<li>Do Track Analysis on the two files. Note that their volumes are only 0.1 dB apart
		<li>Without closing MP3Gain, listen to the mp3 files using your favorite mp3 player. Note how they're approximately the same listening volume
		<li>Now in MP3Gain, do &quot;Modify Gain -&gt; Apply Max Noclip Gain&quot; (or press Ctrl-X). The two files are now peak normalized.
		<li>Listen to the mp3 files again. Even though their maximum amplitudes are now almost exactly the same, song clip 2 now sounds much too loud.
	</ol>
	</p>"),
									array ("lossless","Does normalizing the mp3 degrade its quality?",
	"No. MP3Gain does <i>not</i> decode and re-encode the mp3 to change its volume.
	You can change the volume as many times as you want, and the mp3 will sound just as good (or just as bad!) as it did before you started."))),
					array ("Tags",
							array (
									array ("tagbasics","What do these &quot;tags&quot; in the new Beta version do?",
	"Store analysis and undo information inside the mp3 itself. You no longer need to analyze an mp3 more than once"))),
					array ("Troubleshooting",
							array (
									array ("blackscreen","My screen is going completely black!",
	"If your screen goes completely black for a second when you start MP3Gain and then goes black whenever
	you do any analysis or gain changes, then your DOS settings need to be adjusted on your computer. Here's
	what you do:
		<ol>
		<li>Start MP3Gain</li>
		<li>If you're using version 1.2 or later, make sure that &quot;Options - Tags - Ignore tags&quot; is checked.
			This is just to make sure that the analysis will take longer</li>
		<li>Add a large folder full of big mp3s. Again, this is just so that the analysis will take a long time</li>
		<li>Do &quot;Album Analysis&quot;</li>
		<li>While your screen is black, press and hold the &quot;Alt&quot; key, and press the &quot;Enter&quot; key.
			This should make the DOS screen shrink from full-screen to a normal window</li>
		<li>Use your mouse to right-click in the title bar of the DOS window</li>
		<li>In the pop-up menu that appears, select &quot;Properties&quot;</li>
		<li>Somewhere in the Properties window that appears, there will be an option between &quot;Full screen&quot; and &quot;Window&quot;
			(the exact place in the Properties window varies between Windows 95/98/ME/NT/2000/XP)</li>
		<li>Make sure &quot;Window&quot; is selected, and then press the &quot;OK&quot; button</li>
		<li>Windows should then give you a choice; either &quot;Apply properties to current window only&quot; or
			&quot;Save properties for future windows with same title&quot;. Choose the &quot;Save properties...&quot; option</li>
		</ol>"),
									array("tagprobs","My tags (&quot;Artist&quot;, &quot;Title&quot;, etc.) are not working after using MP3Gain",
	"MP3Gain stores &quot;Analysis&quot; and &quot;Undo&quot; information in special tags inside the mp3 file itself. These tags are in the <a href=\"http://doc.hydrogenaudio.org/wikis/hydrogenaudio/APEv2/wikipage_view\">APEv2</a> format. APEv2 tags are carefully designed to <b>not</b> interfere with other tag formats, such as the popular <a href=\"http://www.id3.org/id3v1.html\">ID3v1</a> format.
<p>
Unfortunately, some mp3 players do not strictly adhere to the ID3v1 standard when reading tags. As a result, when MP3Gain writes its APEv2 tags, these mp3 players might get confused and try to read the MP3Gain tags instead of the regular ID3v1 tags such as &quot;Artist&quot;, &quot;Title&quot;, etc. As a result, the player will show random garbage in these fields.
</p>
<p>
(To be fair, the mp3 players that have this problem are actually probably trying to compensate for data corruption that can occur in mp3s due to bad encoders, incomplete downloads, etc.)
</p>
<p>
If you use MP3Gain and discover that your mp3 player has this problem, then here's what you need to do:
<ul>
<li>Select &quot;Options - Tags - Ignore (do not read or write tags)&quot; from the MP3Gain menu. This will prevent MP3Gain from writing any more tags to your files.</li>
<li>To remove tags that MP3Gain has already written, simply load the affected mp3s into MP3Gain and do &quot;Options - Tags - Remove Tags from files&quot;</li>
</ul>
</p>
<p>
<b>IMPORTANT</b><br />
If you choose the &quot;Options - Tags - Ignore&quot; option, then you will not be able to <i>automatically</i> undo changes made by MP3Gain. You <i>will</i> still be able to undo any changes, but you will have to manually keep track of what changes you make to your files.
</p>"))))
?>
<h1 class="hide">Frequently Asked Questions</h1>
<p>
(There are <b>many</b> more FAQs coming soon. This is just a handful to get me started on this page)</p>
<p>
<?php
	foreach ($faqs as $faq) {
		echo "<b>$faq[0]</b>\r\n	<ul>\r\n";
		foreach ($faq[1] as $item) {
			echo "		<li><a href=\"#$item[0]\">$item[1]</a></li>\r\n";
		}
		echo "	</ul>\r\n";
	}
?>
</p>
<p>
<?php
	foreach ($faqs as $faq) {
		echo "<div class=\"faqSection\">\r\n<h2>$faq[0]</h2>\r\n";
		foreach ($faq[1] as $item) {
			echo "<p>\r\n<h3><a name=\"$item[0]\"><b>$item[1]</b></a></h3>\r\n$item[2]\r\n</p>\r\n";
		}
		echo "</div>\r\n";
	}
?>
</p>
<? include("end.php"); ?>
