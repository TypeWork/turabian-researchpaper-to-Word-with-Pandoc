#!/usr/bin/perl

# pandoc-turabian-formatDocx.pl provides additional formatting of a 
# specified .docx document created by pandoc-turabian.pl.
# Copyright (C) 2018  Omar Abdool
# 
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
# 
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
# 
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
#
# Required support files:
#	$HOME/.pandoc/
#		turabian-style-reference.docx
#
# Command:	pandoc-turabian-formatDocx.pl [FileName] [Optional:endnotes]
# Output:	[FileName].docx
#
# Version:	2018/07/27
# 

use 5.010;
use strict;
use warnings;

use Archive::Zip qw/ :ERROR_CODES :CONSTANTS /;
use File::Basename;

use XML::LibXML qw(:libxml);
use XML::LibXML::XPathContext;


print "\nRunning pandoc-turabian-formatDocx.pl...\n";


# Get file name from command line and use for output file
my ($docxFileName, $docxFilePath, $docxFileExt) = fileparse($ARGV[0], qr/\.[^.]*/);

my $docxFile = "${docxFileName}.docx";


# Construct notice string
sub noticeOut {
	my ($noticeType, $noticeStr, $subNoticeStr) = @_;
 	$subNoticeStr //= '';
	return "\n  ${noticeType}: ${noticeStr}\n    $subNoticeStr\n";
}


# Check for 'endnotes' option from command line
my $endnotesOption = 0;
if ( $ARGV[1] ) {
	if ( $ARGV[1] eq 'endnotes' ) { $endnotesOption = 1; }
}


# Create object and read docx file
my $docx = Archive::Zip->new();

$docx->read( $docxFile ) == AZ_OK
	or die noticeOut("Error", "Could not read \'${docxFile}\'.");


sub ModifyDocumentXml {

	# Extract document.xml file from $docx
	my $xmlFileName = 'document';
	$docx->extractMemberWithoutPaths( "word/${xmlFileName}.xml" ) == AZ_OK
		or die noticeOut("Error", "Could not extract \'${xmlFileName}.xml\'.");

	# Open styles.xml and create temporary modified version
	open(my $xmlOriginalFile, "<${xmlFileName}.xml")
		or die noticeOut("Error", "Could not open \'${xmlFileName}.xml\'.\n");
  	binmode $xmlOriginalFile, ':raw';

	open(my $xmlModFile, ">${xmlFileName}-modified.xml")
		or die noticeOut("Error", "Could not open \'${xmlFileName}-modified.xml\'.\n");
  	binmode $xmlModFile, ':raw';


 	my $dom = XML::LibXML->load_xml(IO => $xmlOriginalFile);

	my $xpc = XML::LibXML::XPathContext->new($dom);
	$xpc->registerNs(w => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main');
	$xpc->registerNs(r => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');


	# Array for printing modification notes at end
	my @printModNotes = ("  Modified word/${xmlFileName}.xml:\n");
	

	# If Subtitle exists, append colon to Title
	if ( $xpc->exists('//w:pStyle[contains(@w:val, "Subtitle")]') ) {
		my $wpTitleNodes = $dom->findnodes('//w:p/w:pPr/w:pStyle[contains(@w:val, "Title")]/../..');
		my $wpTitleLastNode = $wpTitleNodes->get_node($wpTitleNodes->size);
		my $wtTitleNodes = $wpTitleLastNode->findnodes('.//w:t');
		my $wtTitleLastNode = $wtTitleNodes->get_node($wtTitleNodes->size);
		$wtTitleLastNode->appendText(':');

		push(@printModNotes, "    Colon appended to title on title page\n");
	}

	# Modify paragraphs that start with <TitlePageInfo />
	foreach my $wpTPInfoNode ($xpc->findnodes('//w:t[starts-with(text(), "<TitlePageInfo />")]/../..')) {
		# Set w:pStyle to "TitlePageInformation"
		my($wpStyleAttr) = $wpTPInfoNode->findnodes('./w:pPr/w:pStyle/@w:val');
		$wpStyleAttr->setValue('TitlePageInformation') if $wpStyleAttr;

		# Remove <titlepageinfo />
		my($wtTPInfoFirstNodeText) = $wpTPInfoNode->findnodes('.//w:t/text()');
		my $wtTempStr = $wtTPInfoFirstNodeText->to_literal;
		$wtTempStr =~ s/\Q<TitlePageInfo \/>\E//g;
		$wtTPInfoFirstNodeText->setData($wtTempStr);
	}

	# Get header1 reference node
	my($header1Ref) = $xpc->findnodes('//w:headerReference');

	# Add w:footnotePr to initial w:sectPr
	my($wSectPrEnd) = $xpc->findnodes('//w:sectPr');

	my $wFootnotePr = $dom->createElement('w:footnotePr');
	$wSectPrEnd->appendChild($wFootnotePr);

	my $wNumFmt = $dom->createElement('w:numFmt');
	if ( $endnotesOption ) {
		$wNumFmt->{'w:val'} = 'chicago';
		
		my $wNumRestart = $dom->createElement('w:numRestart');
		$wNumRestart->{'w:val'} = 'eachPage';
		$wFootnotePr->appendChild($wNumRestart);
	} else {
		$wNumFmt->{'w:val'} = 'decimal';	
	}
	$wFootnotePr->appendChild($wNumFmt);

	# Build w:sectPr node from initial w:sectPr
	my $wSectPrTemp = $dom->createElement('w:sectPr');

	my $wSectPrTempChild = $wFootnotePr->cloneNode(1);
	$wSectPrTemp->appendChild($wSectPrTempChild);

	my @wSectPrChildNames = ('w:endnotePr', 'w:pgSz', 'w:pgMar', 'w:cols', 'w:docGrid');
	for (my $i = 0; $i < @wSectPrChildNames; ++$i) {
		my($wSectPrEndChild) = $wSectPrEnd->findnodes("./$wSectPrChildNames[$i]");
		$wSectPrTempChild = $wSectPrEndChild->cloneNode(1) if $wSectPrEndChild;
		$wSectPrTemp->appendChild($wSectPrTempChild) if $wSectPrTempChild;
	}

	# Build w:pgNumType node for use with header1
	my $wPgNumType = $dom->createElement('w:pgNumType');
	$wPgNumType->{'w:start'} = '1';

	# If <EndTitlePage />, modify w:p and add section break
	if ( my($wpPrEndTitlePage) = $xpc->findnodes('//w:p/w:r/w:t[text()="<EndTitlePage />"]/../../w:pPr') ) {
		my $wSectPrEndTitlePage = $wSectPrTemp->cloneNode(2);

		# Set w:pStyle to "Normal"
		my($wpStyleAttr) = $wpPrEndTitlePage->findnodes('./w:pStyle/@w:val');
		$wpStyleAttr->setValue('Normal') if $wpStyleAttr;

		# Append w:sectPr
		$wpPrEndTitlePage->appendChild($wSectPrEndTitlePage);

		# Remove <EndTitlePage />
		my $wpEndTitlePage = $wpPrEndTitlePage->parentNode;
		foreach my $wrEndTitlePage ($wpEndTitlePage->findnodes('./w:r')) {
			$wpEndTitlePage->removeChild($wrEndTitlePage);
		}
		
		# Insert pgNumType to start at page 1 for header1
		$wSectPrEnd->appendChild($wPgNumType);

		push(@printModNotes, "    Formatted title page created\n");
	}


	# Modify "starred" sections to use "Heading 1 New Page" style
	foreach my $wpSSectNode ($xpc->findnodes('//w:t[starts-with(text(), "<StarredSection />")]/../..')) {
		# Set w:pStyle to "Heading1NewPage"
		my($wpStyleAttr) = $wpSSectNode->findnodes('./w:pPr/w:pStyle/@w:val');
		$wpStyleAttr->setValue('Heading1NewPage') if $wpStyleAttr;

		# Remove <StarredSection />
		my($wtSSectFirstNodeText) = $wpSSectNode->findnodes('.//w:t/text()');
		my $wtTempStr = $wtSSectFirstNodeText->to_literal;
		$wtTempStr =~ s/\Q<StarredSection \/>\E//g;
		$wtSSectFirstNodeText->setData($wtTempStr);
	}
	push(@printModNotes, "    Starred section heading styles set to \"Heading 1 New Page\"\n");

	# If <SectionEndnotes />, modify w:p and insert section break
	if ( $endnotesOption ) {
		if ( my($wpPrSEndnotes) = $xpc->findnodes('//w:p/w:r/w:t[text()="<SectionEndnotes />"]/../../w:pPr')) {
			my $wSectPrSEndnotes = $wSectPrTemp->cloneNode(2);

			# Set w:pStyle to "Heading1SectionEndnotes"
			my($wpStyleAttr) = $wpPrSEndnotes->findnodes('./w:pStyle/@w:val');
			$wpStyleAttr->setValue('Heading1SectionEndnotes') if $wpStyleAttr;

			# Append w:sectPr
			$wpPrSEndnotes->appendChild($wSectPrSEndnotes);
		
			# Change "<SectionEndnotes />" to "Notes"
			my $sEndnotesHeading = $dom->createElement('w:t');
			$sEndnotesHeading->appendText('Notes');

			my $wpSEndnotes = $wpPrSEndnotes->parentNode;
			foreach my $wrSEndnotes ($wpSEndnotes->findnodes('./w:r')) {
				$wpSEndnotes->removeChild($wrSEndnotes);
			}
			my $wrSEndnotes = $dom->createElement('w:r');
			$wpSEndnotes->appendChild($wrSEndnotes);
			$wrSEndnotes->appendChild($sEndnotesHeading);
				
			# Insert pgNumType to start at page 1 with header1
			$wSectPrEnd->removeChild($wPgNumType);
			$wSectPrSEndnotes->appendChild($wPgNumType);
			$wSectPrEnd->removeChild($header1Ref);
			$wSectPrSEndnotes->appendChild($header1Ref);

			push(@printModNotes, "    Endnotes section created\n");
		}
	}
	
	# Write out to $xmlModFile
	print $xmlModFile $dom->toString;

	# close open files and update $docx
	close($xmlOriginalFile);
	close($xmlModFile);
	$docx->updateMember( "word/${xmlFileName}.xml", "${xmlFileName}-modified.xml");

	print @printModNotes;

	return;
}


sub ModifyFootnotesXml {

	# Extract footnotes.xml file from $docx
	my $xmlFileName = 'footnotes';	
	$docx->extractMemberWithoutPaths( "word/${xmlFileName}.xml" ) == AZ_OK
		or die noticeOut("Error", "Could not extract \'${xmlFileName}.xml\'.");

	# Open styles.xml and create temporary modified version
	open(my $xmlOriginalFile, "<${xmlFileName}.xml")
		or die noticeOut("Error", "Could not open \'${xmlFileName}.xml\'.\n");
  	binmode $xmlOriginalFile, ':raw';

	open(my $xmlModFile, ">${xmlFileName}-modified.xml")
		or die noticeOut("Error", "Could not open \'${xmlFileName}-modified.xml\'.\n");
  	binmode $xmlModFile, ':raw';


 	my $dom = XML::LibXML->load_xml(IO => $xmlOriginalFile);

	my $xpc = XML::LibXML::XPathContext->new($dom);
	$xpc->registerNs(w => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main');


	# Create footnote separator w:p and children
	my $wFootnoteSepP = $dom->createElement('w:p');

	my $wFootnoteSepPPr = $dom->createElement('w:pPr');
	$wFootnoteSepP->appendChild($wFootnoteSepPPr);

	my $wFootnoteSepPStyle = $dom->createElement('w:pStyle');
	$wFootnoteSepPStyle->{'w:val'} = 'FootnoteText';
	$wFootnoteSepPPr->appendChild($wFootnoteSepPStyle);

	my $wFootnoteSepInd = $dom->createElement('w:ind');
	$wFootnoteSepInd->{'w:firstLine'} = '0';
	$wFootnoteSepPPr->appendChild($wFootnoteSepInd);

	my $wFootnoteSepR = $dom->createElement('w:r');
	$wFootnoteSepP->appendChild($wFootnoteSepR);
	
	my $wFootnoteSepSeparator = $dom->createElement('w:separator');
	$wFootnoteSepR->appendChild($wFootnoteSepSeparator);


	# Format footnote separators
	foreach my $wFootnoteSep ($xpc->findnodes('//w:footnote[contains(@w:type, "separator")]')) {
		$wFootnoteSep->removeChildNodes();
		$wFootnoteSep->appendChild($wFootnoteSepP);
	}

	# Format footnote continuation separators
	foreach my $wFootnoteContSep ($xpc->findnodes('//w:footnote[contains(@w:type, "continuationSeparator")]')) {
		$wFootnoteContSep->removeChildNodes();
		my $wFootnoteContSepP = $wFootnoteSepP->cloneNode(2);
		$wFootnoteContSep->appendChild($wFootnoteContSepP);		
	}


	# Write out to $xmlModFile
	print $xmlModFile $dom->toString;

	# close open files and update $docx
	close($xmlOriginalFile);
	close($xmlModFile);
	$docx->updateMember( "word/${xmlFileName}.xml", "${xmlFileName}-modified.xml");

	print "  Modified word/${xmlFileName}.xml: Footnote separators updated\n";

	return;
}


sub ModifyStylesXml {

	# Extract styles.xml file from $docx
	my $xmlFileName = 'styles';	
	$docx->extractMemberWithoutPaths( "word/${xmlFileName}.xml" ) == AZ_OK
		or die noticeOut("Error", "Could not extract \'${xmlFileName}.xml\'.");

	# Open styles.xml and create temporary modified version
	open(my $xmlOriginalFile, "<${xmlFileName}.xml")
		or die noticeOut("Error", "Could not open \'${xmlFileName}.xml\'.\n");
  	binmode $xmlOriginalFile, ':raw';

	open(my $xmlModFile, ">${xmlFileName}-modified.xml")
		or die noticeOut("Error", "Could not open \'${xmlFileName}-modified.xml\'.\n");
  	binmode $xmlModFile, ':raw';


 	my $dom = XML::LibXML->load_xml(IO => $xmlOriginalFile);

	my $xpc = XML::LibXML::XPathContext->new($dom);
	$xpc->registerNs(w => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main');


	# Modify each "w:style" that has a "w:styleId" value ending in "Tok"
	foreach my $wStyle ($xpc->findnodes('//w:style[substring(@w:styleId,string-length(@w:styleId) -string-length("Tok") +1) = "Tok"] ')) {
		# Set style priority to '99'
		my $wuiPriority = $dom->createElement('w:uiPriority');
		$wuiPriority->{'w:val'} = '99';
		$wStyle->appendChild($wuiPriority);

		# Make style 'semiHidden' and 'unhideWhenUsed'
		my $wsemiHidden = $dom->createElement('w:semiHidden');
		$wStyle->appendChild($wsemiHidden);
		my $wunhideWhenUsed = $dom->createElement('w:unhideWhenUsed');
		$wStyle->appendChild($wunhideWhenUsed);
	}

	# Write out to $xmlModFile
	print $xmlModFile $dom->toString;

	# close open files and update $docx
	close($xmlOriginalFile);
	close($xmlModFile);
	$docx->updateMember( "word/${xmlFileName}.xml", "${xmlFileName}-modified.xml");

	print "  Modified word/${xmlFileName}.xml: Styles updated\n";

	return;
}


sub ModifySettingsXml {

	# Extract settings.xml file from $docx
	my $xmlFileName = 'settings';	
	$docx->extractMemberWithoutPaths( "word/${xmlFileName}.xml" ) == AZ_OK
		or die noticeOut("Error", "Could not extract \'${xmlFileName}.xml\'.");

	# Open styles.xml and create temporary modified version
	open(my $xmlOriginalFile, "<${xmlFileName}.xml")
		or die noticeOut("Error", "Could not open \'${xmlFileName}.xml\'.\n");
  	binmode $xmlOriginalFile, ':raw';

	open(my $xmlModFile, ">${xmlFileName}-modified.xml")
		or die noticeOut("Error", "Could not open \'${xmlFileName}-modified.xml\'.\n");
  	binmode $xmlModFile, ':raw';


 	my $dom = XML::LibXML->load_xml(IO => $xmlOriginalFile);

	my $xpc = XML::LibXML::XPathContext->new($dom);
	$xpc->registerNs(w => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main');


	# Endnotes set to "End of section" using w:endnotePr

	# Create w:endnotePr with children
	my $wEndnotePr = $dom->createElement('w:endnotePr');

	my $wPos = $dom->createElement('w:pos');
	$wPos->{'w:val'} = 'sectEnd';
	$wEndnotePr->appendChild($wPos);

	my $wNumFmt = $dom->createElement('w:numFmt');
	$wNumFmt->{'w:val'} = 'decimal';
	$wEndnotePr->appendChild($wNumFmt);
	
	# Append Place w:endnotePr after w:footnotePr
	my($wSettings) = $xpc->findnodes('/w:settings');
	my($wFootnotePr) = $wSettings->findnodes('./w:footnotePr');
	$wSettings->insertAfter( $wEndnotePr, $wFootnotePr );	


	# Write out to $xmlModFile
	print $xmlModFile $dom->toString;

	# close open files and update $docx
	close($xmlOriginalFile);
	close($xmlModFile);
	$docx->updateMember( "word/${xmlFileName}.xml", "${xmlFileName}-modified.xml");

	print "  Modified word/${xmlFileName}.xml: Endnotes set to \"End of section\"\n";

	return;
}


# Modify .xml files in .docx
ModifyDocumentXml();
ModifyFootnotesXml();
ModifyStylesXml();

# Modify endnotes settings in settings.xml
if ( $endnotesOption ) {
	ModifySettingsXml();
	print noticeOut("Notice", "Use the \"PandocFormatEndnotes\" macro to complete formatting endnotes.");
}


# Write updated file to $docx
$docx->overwrite() == AZ_OK
	or die noticeOut("Error", "Could not modify \'${docxFile}\'.");


# Remove/delete temporary files
unlink 'document.xml', 'document-modified.xml';
unlink 'footnotes.xml', 'footnotes-modified.xml';
unlink 'styles.xml', 'styles-modified.xml';

if ( $endnotesOption ) { unlink 'settings.xml', 'settings-modified.xml'; }

print "  Done modifying ${docxFile}\n";

