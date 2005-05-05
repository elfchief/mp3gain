<?php 
  include('pagebase.php');
  //initialize variables
    if (!(isset($linkHere))) {
        $linkHere = basename($_SERVER['PHP_SELF']);
    }
  
  $menuPages = array ( 
					array ( "first" => true, "page" => "index.php",			"title" => "Home",			"desc" => "Welcome screen, with latest few news items" ), 
  					array ( "page" => "news.php",			"title" => "News",			"desc" => "What's new in the world of MP3Gain" ),
  					array ( "page" => "download.php",		"title" => "Downloads",		"desc" => "Download MP3Gain in its various forms, including source code" ),
                    array ( "page" => "translation.php",	"title" => "Translations",	"desc" => "See what translations of MP3Gain are available, and how to make your own"),
                    array ( "page" => "faq.php",			"title" => "FAQ",			"desc" => "Frequently Asked Questions (and answers!)")
			   );
?>
<script language="javascript" type="text/javascript">
    function fn_setMenuText(asText) {
        menuText.innerHTML = asText;
    }
</script>
<div id="header">
<a href="index.php"><img src="images/mp3gainlogosmall.gif" alt="MP3Gain" WIDTH="64" HEIGHT="64" border="0" align="center"><img src="images/mp3gainlogo.gif" alt="MP3Gain" align="center" width="223" height="55" border="0"></a>
</div>
<div id="menu">
<ul>
<?php
foreach ($menuPages as $mPage) {
	if ($mPage['page'] != $linkHere) {
		echo _descText($mPage,true);
		if (isset($mPage['sub'])) {
			foreach (($mPage['sub']) as $sPage) {
				if ($sPage['page'] == $linkHere) { //show sub-menu, mark selected
					_doSubMenu($mPage['sub'],$linkHere);
					break;
				}
			}
		}
	} else {
		echo _descText($mPage,false);
		if (isset($mPage['sub'])) {
			_doSubMenu($mPage['sub'],"");
		}
	} 
}
?>

</ul>
</div>
<div id="content">
<?php

    function _doSubMenu($arrSubs, $link) {
        echo "\n" . '					<table width="100%" border="0"><tr><td align="right"><table border="0" cellspacing="0" summary="submenu">';
        foreach ($arrSubs as $subPage) {
            echo "\n" . '						<tr><td>&nbsp&nbsp&nbsp;&nbsp&nbsp;<b>-</b></td>';
            echo '<td align="left">';
            if ($subPage[0] != $link) {
               echo _descText($subPage,true);
            } else {
               echo _descText($subPage,false);
            }
            echo '</td></tr>';
       }
       echo "\n" . '					</table></td></tr></table>';
    }
    
    function _descText($arr,$isHref) {
        //$ret = "onmouseover=\"fn_setMenuText('" . $arr['desc'] . "')\" onmouseout=\"fn_setMenuText('&nbsp;')\"";
		$ret = "title=\"" . $arr['desc'] . "\"";
		if (isset($arr['first'])) {
			$ret .= " id=\"leftie\"";
		}
		$pre = "<li>";
		$post = "</li>";
        if ($isHref) {
            return "$pre<a href=\"" . $arr['page'] . "\" $ret>" . $arr['title'] . "</a>$post";
        } else {
            return "$pre<span $ret>" . $arr['title'] . "</span>$post";
        }
    }
?>
