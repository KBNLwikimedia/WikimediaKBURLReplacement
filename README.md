**Code, scripts and stories about replacing outdated or non-persistent URLs of KB services in Wikipedia, Commons and Wikidata**

Read this [article](stories/Making%20references%20to%20Dutch%20newspapers%20in%20Wikipedia%20more%20sustainable.md) to understand why and how the KB replaces outdates URLs in Wikipedia.

* Folder *[ScriptsMerlijnVanDeen](ScriptsMerlijnVanDeen)*
  - The technique for replacing URLs is detailed in [this blogpost](https://web.archive.org/web/20200522204706/https://merlijn.vandeen.nl/2015/kb-replace-dead-links.html) by Merlijn van Deen. See the 4 screenshot below.
   - The Python scripts and [PAWS](https://wikitech.wikimedia.org/wiki/PAWS) commands are available [from this folder](ScriptsMerlijnVanDeen/scripts)
<img src="stories/images/blogMvD_part1.jpg" align="left" width="200"/><img src="stories/images/blogMvD_part2.jpg" align="left" width="200"/><img src="stories/images/blogMvD_part3.jpg" align="left" width="200"/><img src="stories/images/blogMvD_part4.jpg" align="left" width="200"/>
 
* Folder *[GvN](GvN)* : Pywikibot commands for [PAWS](https://wikitech.wikimedia.org/wiki/PAWS) to replace links to Geheugen van Nederland (GvN) in Wikipedia based on [this example](https://www.mediawiki.org/wiki/Manual:Pywikibot/PAWS#A_real_script_example)

* Folder *[ScriptsHayKranen_KrantenKB](ScriptsHayKranen_KrantenKB)* : two Pyton scripts to replace URLs of newspapers in Delpher. Not yet workked with.

A Jypyter notebooks ([PAWS](https://wikitech.wikimedia.org/wiki/PAWS)) implementation of these scripts is available from https://paws-public.wmflabs.org/paws-public/User:OlafJanssenBot/WikipediaURLReplacement/
