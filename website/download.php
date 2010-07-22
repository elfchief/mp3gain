<?php

$pageTitle = "MP3Gain Downloads";

include("start.php");

$Gx = 1;
$Gy = 3;
$Gz = 4;

$Gsx = 1;
$Gsy = 2;
$Gsz = 5;

$Cx = 1;
$Cy = 5;
$Cz = 1;

$downloads = array ( "MP3Gain-Windows %28Stable%29/$Gsx.$Gsy.$Gsz" => array ("MP3Gain-Windows (Stable)", array (
							array ("mp3gain-win-".$Gsx."_".$Gsy."_".$Gsz.".exe","Normal MP3Gain install for version $Gsx.$Gsy.$Gsz<br>
			<b>This is what most people will want to download.</b>"),
							array ("mp3gain-win-".$Gsx."_".$Gsy."_".$Gsz.".zip","Normal MP3Gain, but with no installer"),
							array ("mp3gain-win-full-".$Gsx."_".$Gsy."_".$Gsz.".exe","Exactly the same as the Normal install, but also includes the Microsoft Visual Basic run-time files.
			The VB run-time files only need to be installed on a computer <i>once</i>, so they might already be in your Windows folder.
			If you're not sure, then go ahead and download this Full version. Or if you want to save some download time, then try
			the Normal install first.<p>
			If you ever download a newer version of MP3Gain after doing a Full install, you will only need the Normal version.</p>"),
							array ("mp3gain-win-full-".$Gsx."_".$Gsy."_".$Gsz.".zip","Full MP3Gain (i.e. Normal MP3Gain plus VB run-time files), but with no installer"),
							array ("mp3gain-win-gui-".$Gsx."_".$Gsy."_".$Gsz."-src.zip","Visual Basic source files used to create the MP3Gain GUI.
		     The GUI is just a front end for the command-line program mp3gain.exe, so if you want to completely re-compile MP3Gain from scratch, then
		     you'll also need the <a href=\"http://prdownloads.sourceforge.net/mp3gain/mp3gain-".$Cx."_".$Cy."_".$Cz."-src.zip?download\">mp3gain-".$Cx."_".$Cy."_".$Cz."-src.zip</a> file."))),
					"MP3Gain-Windows %28Beta%29/$Gx.$Gy.$Gz" => array ("MP3Gain-Windows (Beta)", array (
							array ("mp3gain-win-".$Gx."_".$Gy."_".$Gz.".exe","Normal MP3Gain install for version $Gx.$Gy.$Gz <b>Do not</b> use this version unless you really need the experimental Unicode support. There seem to be some cases where this version accidentally shortens the filename. I'm still figuring it out."),
							array ("mp3gain-win-".$Gx."_".$Gy."_".$Gz.".zip","Normal MP3Gain, but with no installer"),
							array ("mp3gain-win-full-".$Gx."_".$Gy."_".$Gz.".exe","Exactly the same as the Normal install, but also includes the Microsoft Visual Basic run-time files.
			The VB run-time files only need to be installed on a computer <i>once</i>, so they might already be in your Windows folder.
			If you're not sure, then go ahead and download this Full version. Or if you want to save some download time, then try
			the Normal install first.<p>
			If you ever download a newer version of MP3Gain after doing a Full install, you will only need the Normal version.</p>"),
							array ("mp3gain-win-full-".$Gx."_".$Gy."_".$Gz.".zip","Full MP3Gain (i.e. Normal MP3Gain plus VB run-time files), but with no installer"),
							array ("mp3gain-win-gui-".$Gx."_".$Gy."_".$Gz."-src.zip","Visual Basic source files used to create the MP3Gain GUI.
		     The GUI is just a front end for the command-line program mp3gain.exe, so if you want to completely re-compile MP3Gain from scratch, then
		     you'll also need the <a href=\"http://prdownloads.sourceforge.net/mp3gain/mp3gain-".$Cx."_".$Cy."_".$Cz."-src.zip?download\">mp3gain-".$Cx."_".$Cy."_".$Cz."-src.zip</a> file."))),
					"mp3gain/$Cx.$Cy.$Cz" => array ("mp3gain (command-line back end)", array (
							array ("mp3gain-".$Cx."_".$Cy."_".$Cz."-src.zip","C++ files (plus Visual C++ project information files) used to create the mp3gain.exe back end"),
							array ("mp3gain-dos-".$Cx."_".$Cy."_".$Cz.".zip","Command-line only version of mp3gain. If you download any of the Windows MP3Gain files above, this file is included."))));
?>
<h1 class="hide">Downloads</h1>
<p><strong>AACGain support</strong>: You will also need to <a href="http://www.rarewares.org/aac.html">download AACGain</a>, rename it to "mp3gain.exe", and put it in the MP3Gain folder after installation.
<p>Here's a list of what you'll find at the SourceForge <a href="https://sourceforge.net/project/showfiles.php?group_id=49979">download page for MP3Gain</a>.<br>
The <a href="http://homepage.mac.com/beryrinaldo/AudioTron/MacMP3Gain/">MacMP3Gain page</a> has information about a Macintosh version of MP3Gain.
<p>
<table class="tableBorder">
<?php
	foreach($downloads as $sectionUrl => $section) { ?>
	<tr>
		<th colspan="2" bgcolor="Silver" class="thBorder"><?php echo $section[0]?></th>
	<tr>
<?php
		foreach ($section[1] as $item) { ?>
	<tr>
		<td nowrap class="tdBorder">
			<a href="https://sourceforge.net/projects/mp3gain/files/<?php echo $sectionUrl ?>/<?php echo $item[0]?>/download"><?php echo $item[0]?></a></td>
		<td class="tdBorder"><?php echo $item[1]?></td>
	</tr>
<?php
		}
	}
?>
</table>
</p>
<? include("end.php"); ?>
