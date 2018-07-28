#!/usr/bin/perl

# pandoc-turabian-prepTeX.pl creates a .tex document, modified from a 
# specified LaTeX document that uses the "turabian-researchpaper" 
# document class, for use with pandoc-turabian.pl.
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
#		turabian-latex-preamble.tex
# 
# Command:	pandoc-turabian-prepTeX.pl [FileName] [Optional:endnotes]
# Output:	[FileName]-pandoc.tex
# 	Output if "filecontents" environment for .bib: [BibFileName.bib]
# 
# Version: 2018/07/27
# 

use strict;
use utf8;
use File::Basename;
use DateTime;


print "\nRunning pandoc-turabian-prepTeX.pl...\n";


# Get file name from command line and use for output file
my ($texFileName, $texFilePath, $texFileExt) = fileparse($ARGV[0], qr/\.[^.]*/);

my $outputFileName = "${texFileName}-pandoc.tex";

# Construct notice string
sub noticeOut {
	my ($noticeType, $noticeStr, $subNoticeStr) = @_;
 	$subNoticeStr //= '';
	return "\n  ${noticeType}: ${noticeStr}\n    $subNoticeStr\n";
}

open(my $originalFile, "<${texFileName}.tex")
	or die noticeOut("Error", "Could not open '${texFileName}.tex'.");
open(my $outputFile, ">${outputFileName}")
	or die noticeOut("Error", "Could not open '${outputFileName}'.");

# Check and use UTF8 encoding with files
binmode $originalFile, ":encoding(UTF-8)";
binmode $outputFile, ":encoding(UTF-8)";


# Check for 'endnotes' option from command line
my $endnotesOption = 0;
if ( $ARGV[1] ) {
	if ( $ARGV[1] eq 'endnotes' ) { $endnotesOption = 1; }
}


# Variables for finding and extracting .bib data
my $bibFileName = 'bibFileNameTemp';

my $inPreamble = 1;
my $parseFileContents = 0;

my @bibFileContent = ();
my $bibFileLineNum = 0;


# If turabian-latex-preamble.tex exists, include in the .tex file preamble
my $inclPreambleFileName = "turabian-latex-preamble.tex";
my $inclPreambleRelDir = ".pandoc";
my $inclPreambleLoc = $ENV{"HOME"} . "/${inclPreambleRelDir}/${inclPreambleFileName}";

my $inclPreambleCmd = '';

if ( -e "$inclPreambleLoc" ) {
	print "  Including in preamble: ~/${inclPreambleRelDir}/${inclPreambleFileName}\n";
	$inclPreambleCmd = "\\include{${inclPreambleLoc}}";
}
else {
	print "  File not found: ~/${inclPreambleRelDir}/${inclPreambleFileName}\n";
}


# Get today's date as string
my $todayDate = DateTime->today(time_zone => 'local')->strftime('%B %e, %Y');

# Arrays for find/replace modifications
my @textModStrings = (
	[ '{turabian-researchpaper}', "{turabian-researchpaper}\n\n${inclPreambleCmd}" ],
	[ '"', '\'\'' ],
	[ '\today', $todayDate ],
	[ '\section*{', '\section*{<StarredSection />' ]
);

# Preserve endnotes with additional markup
my $numEndnotes = 0;
if ( $endnotesOption ) {
	push(@textModStrings, [ '\endnote{', '\footnote{<Endnote /> ' ]);
	push(@textModStrings, [ '\theendnotes', '<SectionEndnotes />' ]);
} else {
	push(@textModStrings, [ '\endnote{', '\footnote{' ]);
}

# Parse through $originalFile with write to $outputFile
while (<$originalFile>) {
	# If filecontents environment for .bib, extract data to @bibFileContent
	if ( ($inPreamble == 1) && ($parseFileContents < 2) ) {
		if ( $_ =~ m/\Q\begin{filecontents}{\E([^.]*).bib\Q}\E/ ) {
			$bibFileName = $1;
			if ( $bibFileName eq '\jobname' ) {
				$bibFileName = $texFileName;
			}
			$_ = '';
			$parseFileContents = 1;
		} elsif ( $parseFileContents == 1 ) {
			if ( $_ =~ m/\Q\end{filecontents}\E/ ) {
				$parseFileContents = 2;
			} else {
				$bibFileContent[$bibFileLineNum] = $_;
				$bibFileLineNum++;
			}
			$_ = '';
		} elsif ( $_ =~ m/\Q\begin{document}\E/ ) {
			$inPreamble = 0;
		}
	}
	# Count number of '\endnote'
	if ( $_ =~ m/\Q\endnote\E/ ) { ++$numEndnotes; }

	# Implement find/replace modifications
	for my $i ( 0 .. @textModStrings) {
		$_ =~ s/\Q$textModStrings[$i][0]\E/$textModStrings[$i][1]/g;
	}
	# Copy modified lines to $outputFile
	print $outputFile "$_";
}

if ( $numEndnotes > 0 ) {
	my $plS ='';
	my $plA ='a ';
	if ( $numEndnotes > 1 ) { $plS = 's'; $plA = ''; }

	my $subNotice = '';
	if ( $endnotesOption ) {
		$subNotice = "  Endnotes, preserved as footnotes, start with \'<Endnote />\' tags.\n";
	}
	print noticeOut(
		"Warning",
		"${numEndnotes} endnote${plS} found and preserved as ${plA}footnote${plS}.",
		$subNotice);
}


# If filecontents environment for .bib, write out @bibFileContent into new .bib file
if ( $parseFileContents == 2 ) {
	open(my $outputBibFile, ">${bibFileName}.bib")
		or die noticeOut("Error", "Could not open '${bibFileName}.bib'.");
	binmode $outputBibFile, ":encoding(UTF-8)";

	for (my $i = 0; $i < @bibFileContent; ++$i) {
		print $outputBibFile $bibFileContent[$i];
	}
	
	print "  Created resource file: ${bibFileName}.bib\n";
	close($outputBibFile);
}

print "  Created temporary file for Pandoc: $outputFileName\n";

close($originalFile);
close($outputFile);

