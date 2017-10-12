#!/usr/bin/perl -w

#
# Perl routine to convert a .jpz crossword to a Postscript/PDF
# 2011 Alex Boisvert and others
# Usage: jpz2pdf.pl file.jpz
# Options: -b for black squares
# -u for grid in upperright
# -p for output to printer

use strict;

use XML::Simple;
{ # Ugly fix to make it work with pp
	no strict 'refs';
	*{"XML::SAX::"}{HASH}{parsers} = sub {
		return [ {
			'Features' => {
				'http://xml.org/sax/features/namespaces/' => '1'
			},
			'Name' => 'XML::SAX::PurePerl'
		}
		]
	};
}
use XML::SAX::PurePerl;
use IO::Uncompress::Unzip qw(unzip $UnzipError) ;
use Encode;
use HTML::Entities;

##
#use Data::Dumper;
##

my $usage = <<EOF;
USAGE: $0 [options] <file1.jpz> [file2.jpz ...]

Options: -p: output directly to printer instead of PDF
	 -b: black squares are black (default: gray)
	 -u: grid in upper right (default: lower right)
EOF

# yay global variables
my $psmode = 1;
my $printmode = 0;
my $gray = 0.75; # spectrum from 0 = black to 1 = white
my $gridup = 0; # grid goes on bottom by default
my $diagramless = 0;

# process options
while ($ARGV[0] =~ /^-/) {
    if ($ARGV[0] eq '-p') { $printmode = 1; }
    elsif ($ARGV[0] eq '-b') { $gray = 0.0; }
    elsif ($ARGV[0] eq '-u') { $gridup = 1; }
    else { print "Ignoring unrecognized option $ARGV[0]\n"; }
    shift;
}

# locate the .jpz file(s)
die $usage unless $ARGV[0];

while (my $infile = shift) {
	my %puz_hash = parse_jpz($infile);
	
	#print Dumper(\%puz_hash);
	#exit(0);

	my $outfile = $infile;
	$outfile =~ s/\..*?$/\.ps/;
	
    unless (open OUT, ">$outfile") {
		print "Can't write to $outfile\n";
		next;
    }

	my $data = psify_puzzle(\%puz_hash);
	close OUT;
	
	print "$infile -> ";
	if ($printmode) {
		system("lpr $outfile");
		if ($? == 0) { # if it worked, clean up by deleting the .ps file
			unlink($outfile);
			print "-> printer\n";
		} else {
			print "print failed!\n";
		}
	} else {
		system("ps2pdf $outfile");
		if ($? == 0) { # if it worked, clean up by deleting the .ps file
			unlink($outfile);
			$outfile =~ s/ps$/pdf/;
			print "-> $outfile\n";
		} else {
			print "ps to pdf conversion failed!\n";
		}
	} # $printmode
} # end while
		

##
# SUBS
##

sub psify_puzzle {
  my $pz = shift;
  my %puz_hash = %$pz;
  
  my @grid = @{$pz->{'puzzle'}};
  
  my $w = $#{$grid[0]}+1;
  my $h = $#grid+1;
  
  my $page_height = 11.0;
  my $page_width = 8.5;
  print OUT <<EOPS;
%!PS-Adobe-2.0
%%PageOrder: Ascend
%%Title: $pz->{'title'}
%%Creator: jpz2pdf.pl (C) 2007-11 Beth Skwarecki, Richard M Kreuter, Alex Boisvert, Joon Pahk
%%BoundingBox: 0 0 612 792
%%DocumentPaperSizes: Letter
%%EndComments
%%BeginProlog

% These are all the various global parameters to this PS program.

% Units of measure.
/cm { 72 2.54 div mul } bind def
/in { 72 mul } bind def

% Physical dimensions of the page.
/page-width { $page_width in } bind def
/page-height { $page_height in } bind def

% Logical dimensions of the puzzle grid
/grid-rows $h def
/grid-cols $w def

% Font and size to use for the header text
/title-font-name /ISOArial-Bold def
/title-font-size 16 def
/author-font-name /ISOArial def
/author-font-size 14 def
/copyright-font-name /ISOArial def
/copyright-font-size 9 def

% Font and size to use for the clue subheader text.  Note: these fonts
% must be ISO-8859-1 encoded.  See the procedure RE below.  Note also:
% Unicode-aware emacsen may transcode iso-8859-1 encoded characters to
% Unicode, which will screw things up.
/clue-header-font-name /ISOArial-Bold def
/clue-header-font-size 12 def

% Font and size to use for the clue text.
/clue-font-name /ISOArial def
/clue-font-size grid-rows 17 le { 11 } {9} ifelse def

% Font and size to use for the labels in the puzzle boxes
/number-font-name /ISOArial def
/number-font-size grid-rows 17 le { 8 } { 6 } ifelse def

% Vertical and horizontal margins around the content in the page.
/page-margin-top { .50 in } bind def
/page-margin-bottom { .50 in } bind def
/page-margin-left { .50 in } bind def
/page-margin-right { .50 in } bind def

% Spacing between columns of clues.
/column-space { 1 6 div in } bind def

% How much of the first page to devote to the puzzle grid.
/grid-share { 7 10 div } bind def
/numcols 15 grid-rows le { 4 } { 3 } ifelse def

% The rest are procedure definitions and global variables,
% should not need editing

% Re-encode fonts for ISO8859-1.
/RE { % /NewFontName [NewEncodingArray] /FontName RE -
   findfont dup length dict begin
   {
       1 index /FID ne
       {def} {pop pop} ifelse
   } forall
   /Encoding exch def
   /FontName 1 index def
   currentdict definefont pop
   end
} bind def
/ISOArial ISOLatin1Encoding /Arial RE
/ISOArial-Bold ISOLatin1Encoding /Arial-Bold RE

% usable width of page
/page-visible-horizontal {
    page-width page-margin-left page-margin-right add sub
} bind def

% width of a single column
/column-width {
    page-visible-horizontal numcols 1 sub column-space mul sub numcols div
} bind def

% Total size of the puzzle grid
/grid-size page-visible-horizontal grid-share mul def

% x position of the left edge of the grid
/grid-horizontal-position {
    page-width grid-size page-margin-right add sub 
} bind def

EOPS

    if ($gridup) { # grid on top
	print OUT <<EOPS;
% y position of the bottom of the grid
/grid-vertical-position {
    page-height
    page-margin-top grid-size title-font-size author-font-size 2 mul
    add add add
    sub
} bind def

% y position of the top of the clue column
/clue-top {
    grid-top
    first-page {
	column-start column-width add grid-horizontal-position ge {
	    pop
	    grid-vertical-position
	    copyright-font-size sub
	    column-space sub
	} if
    } if
} bind def

% calculates the height at which text must be wrapped
/clue-bottom {
    page-margin-bottom
} def
EOPS
    } else { # grid on bottom
	print OUT <<EOPS;
% y position of the bottom of the grid
/grid-vertical-position {
    page-margin-bottom copyright-font-size add
} bind def

% y position of the top of the clue column
/clue-top {
    page-height
    page-margin-top title-font-size author-font-size 2 mul
    add add
    sub
} bind def

% calculates the height at which text must be wrapped
/clue-bottom {
    page-margin-bottom
    first-page {
	column-start column-width add grid-horizontal-position ge {
	    pop
	    grid-top
	} if
    } if
} def
EOPS
    }

    print OUT <<EOPS;

% y position of the top of the grid
/grid-top {
    grid-vertical-position grid-size add
} bind def

% Physical size of each puzzle grid cell.
% Slightly modified from the original to accomodate nonsquare crosswords --Alex
/cell-size {
    grid-cols grid-rows le {
	grid-size grid-rows div
    } {
	grid-size grid-cols div
    } ifelse
} bind def

% Radius for circles, if needed.  It's half the cell size.
/cell-radius {
    cell-size 2 div
} bind def

% Make a box path.
/box {        % w h
  dup         % w h h
  0 exch      % w h 0 h
  rlineto     % w h
  exch        % h w
  dup 0       % h w w 0
  rlineto     % h w
  exch        % w h
  -1 mul      % w -h
  0 exch      % w 0 -h
  rlineto     % w
  -1 mul 0    % -w 0
  rlineto
} def

% Make a square box path.
/square-box { % w
  dup box
} bind def

% Draw a horizontal row of square boxes.
/draw-grid-row { % cols %% This was wrong in original version --Alex
  1 sub
  0 1 3 -1 roll {
    gsave
      cell-size mul 0 translate
      newpath 0 0 moveto
      cell-size square-box
      $gray setgray
      stroke
    grestore
  } for
} def

% Draw a rectangle of rows of boxes.
/draw-grid { % cols rows %% This has been switched from the original --Alex
  1 sub
  0 1 3 -1 roll {
    gsave
      cell-size mul 0 exch translate
      dup draw-grid-row
    grestore
  } for
  pop
} def

% Move the current path to the bottom left corner of the cell at (row,
% column). Note: row/column is really y/x; inverted to match row-major
% ordering in the host language.
/moveto-cell { % row col
  exch 1 add exch
  cell-size mul exch
  grid-rows exch sub
  cell-size mul
  moveto
} def

% Fill the grid cell at (row, col).
/fill-cell { % row col
  newpath
  moveto-cell
  cell-size square-box
  $gray setgray			% fill it in gray
  fill
  0.0 setgray			% back to black
} def

% Draw a circle in the square at (row, col). -- Alex
/circle-cell { %row col
  exch 1 add exch		% (row+1) (col)
  cell-size mul			% (row+1) (col*cell-size)
  cell-radius add		% (row+1) (col*cell-size+cell-radius)
  exch					% (col*cell-size+cell-radius) (row+1)
  grid-rows exch sub	% (col*cell-size+cell-radius) (grid-rows - row - 1)
  cell-size mul			% This is getting too complicated.  It moves to 
  cell-radius add 		% the center of the desired square.
  cell-radius 0 360 arc closepath	% Draw the circle
  $gray setgray			% circles are gray
  stroke
  0.0 setgray			% back to black
} def

% Label the grid cell at (row, col) with a string.
/number-cell { % string row col
  newpath
  number-font-name findfont number-font-size scalefont setfont
  moveto-cell
  0 cell-size rmoveto
  2 -1 number-font-size mul rmoveto
  show fill
} def

%% Stuff for filling text in columns.

% column number (leftmost column on first page = 0)
/column-number 0 def
% are we still on the first page (i.e. the page with the grid)?
/first-page true def

% gives the x coordinate of the current column's left edge
/column-start {
    column-number numcols mod dup
    column-width mul exch
    column-space mul add
    page-margin-left add
} def

% gives the current indentation of the column
/column-indent 0 def

% moves to the next column
/nextcolumn {					% x y
    % bookkeeping
    /column-number column-number 1 add def
    column-number numcols eq {
	showpage
	/first-page false def
    } if
    column-start clue-top moveto
    pop pop currentpoint			% x y
} def

% calculates the height of a wrapped line of text in the current font in
%	the current column
/text-height {					% text
    stringwidth pop				% textwidth
    column-width div cvi 1 add			% numlines
    curr-font-size mul				% textheight
} def

% goes to the next column if the text won't fit in the current column
/nextcolumn-maybe {				% x y textheight
    1 index exch sub				% x y y-textheight
    clue-bottom lt {				% are we too low?
	nextcolumn				% x y
    } if
} def

% moves to the next line
/nextline {
    currentpoint pop column-start sub -1 mul curr-font-size -1 mul rmoveto
    column-indent 0 rmoveto
} def

% tests to see if we should move to the next line (wordwrap), and does so
/nextline-maybe {   % colwidth text
  2 copy            % colwidth text colwidth text
  stringwidth pop   % colwidth text colwidth textwidth
  currentpoint pop  % colwidth text colwidth textwidth hpos
  add               % colwidth text colwidth hoffset
  exch              % colwidth text hoffset colwidth
  column-start add  % colwidth text hoffset colright
  gt {              % colwidth text
    nextline
  } if
} def

% recursive function to show a line of text, wrapping when necessary
/showline {				% colwidth text
    dup length 0 gt { 			% colwidth text
	( ) search
	{				% colwidth text2 ( ) text1
            4 -1 roll dup 5 1 roll exch	% colwidth text2 ( ) colwidth text1
	    nextline-maybe		% colwidth text2 ( ) colwidth text1
            show pop			% colwidth text2 ( )
            show			% colwidth text2
            showline
	}
	{				% colwidth text
	    nextline-maybe		% colwidth text
	    show pop			%
	} ifelse
    } if
} def

% helper function used for show-clue and show-clue-header
/show-text-in-column-at { % colwidth text x y
  moveto
  showline
} def

% Show the title
/show-title { % x y text
  title-font-name findfont title-font-size scalefont setfont
  /curr-font-size title-font-size def
  newpath
  3 1 roll moveto
  show
  nextline
  currentpoint
  fill
} def

% Show the author
/show-author { % x y text
  author-font-name findfont author-font-size scalefont setfont
  /curr-font-size author-font-size def
  newpath
  3 1 roll moveto
  show
  nextline
  nextline
  currentpoint
  fill
} def

% Show the copyright
/show-copyright { % text
  copyright-font-name findfont copyright-font-size scalefont setfont
  /curr-font-size copyright-font-size def
  newpath
  grid-horizontal-position grid-vertical-position copyright-font-size sub moveto
  show
  nextline
  fill
} def

% Indent text in a column, used for show-clue
/indent-text {				% x y (# clue)
  dup ( ) search pop			% x y (# clue) clue ( ) #
  3 1 roll pop pop			% x y (# clue) #
  stringwidth pop ( ) stringwidth pop add /column-indent exch def
  newpath
  3 1 roll moveto
  column-width exch currentpoint show-text-in-column-at
  /column-indent 0 def
  nextline
  currentpoint				% put x y back on the stack
  fill
} def

% Show the clue header ("Across" or "Down") and then the first clue
% 	(they're married together for widow protection)
/show-first-clue {			% x y clue header
  % first check to see if it fits in the current column
  4 2 roll 3 index			% clue header x y clue
  /curr-font-size clue-font-size def
  text-height				% clue header x y clueheight
  clue-header-font-size add		% clue header x y totalheight
  nextcolumn-maybe			% clue header x y
  % now show the header
  clue-header-font-name findfont clue-header-font-size scalefont setfont
  /curr-font-size clue-header-font-size def
  newpath
  moveto				% clue header
  column-width exch currentpoint show-text-in-column-at
  nextline
  currentpoint				% clue x y
  fill
  % and now the clue
  clue-font-name findfont clue-font-size scalefont setfont
  /curr-font-size clue-font-size def
  3 2 roll indent-text			% x y
} def

% Show the clue.
/show-clue {				% x y (# clue)
  clue-font-name findfont clue-font-size scalefont setfont
  /curr-font-size clue-font-size def
  % first, check to see if it'll fit in the column
  dup 4 1 roll				% text x y text
  text-height				% text x y textheight
  nextcolumn-maybe			% text x y
  3 2 roll				% x y text
  % then show it
  indent-text				% x y
} def
%%EndProlog

EOPS

    print OUT <<EOPS;
gsave
  grid-horizontal-position grid-vertical-position translate
  newpath
  0 0 moveto
  grid-cols grid-rows draw-grid	%% Modified from the original --Alex
EOPS
  if (!$diagramless) { # can leave the grid blank for a diagramless
   for my $i (0..$h-1) {
	for my $j (0..$w-1) {
	  my $grid_num;
	  if (ref($grid[$i][$j]) eq 'HASH') {
		$grid_num = $grid[$i][$j]{'cell'};
		if ($grid[$i][$j]{'style'}->{'shapebg'} eq 'circle') {printf OUT "  %d %d circle-cell\n", $i,$j;}
	  }
	  else {$grid_num = $grid[$i][$j];}
	  if ($grid_num eq '#') {
	    # black square
	    printf OUT "  %d %d fill-cell\n", $i, $j;
	  } elsif ($grid_num > 0) {
	    # white square with number
	    printf OUT "  (%d) %d %d number-cell\n", $grid_num, $i, $j;
	  } else {
	    # white square without number, ignore
	  }
	}
      }
  }
  
  print OUT <<EOPS;
grestore

gsave
  newpath
  % put the initial x, y on the stack
  page-margin-left page-height page-margin-top sub
EOPS

my @across = @{$puz_hash{'clues'}->{'Across'}};
my @down = @{$puz_hash{'clues'}->{'Down'}};

  printf OUT "  (%s) show-title\n", $pz->{'title'};
  printf OUT "  (%s) show-author\n", $pz->{'author'};
  printf OUT "  (%s) show-copyright\n", $pz->{'copyright'};
# printf OUT "  (ACROSS) show-clue-header\n";
  my $fa = shift @across;
  printf OUT "  (%d. %s) (ACROSS) show-first-clue\n", $fa->[0], $fa->[1];
  for my $aref (@across) {
    printf OUT "  (%d. %s) show-clue\n", $aref->[0], $aref->[1];
  }

  printf OUT "  ( ) show-clue\n";	# Blank space between sections

# printf OUT "  (DOWN) show-clue-header\n";
  my $fd = shift @down;
  printf OUT "  (%d. %s) (DOWN) show-first-clue\n", $fd->[0], $fd->[1];
  for my $dref (@down) {
    printf OUT "  (%d. %s) show-clue\n", $dref->[0], $dref->[1];
  }

  print OUT <<EOPS;
  pop pop
  fill
grestore
showpage
EOPS
}

sub parse_jpz
{
	my $name = shift;
	
	# read in the .jpz file
	my $data1;
	open IN, $name or die "Can't read $name";
	{
		local $/;
		$data1 = <IN>;
	}
	close IN;

	my $z;

	# Unzip the data if it's zipped
	unless ($data1 =~ /crossword-compiler/) {
		my $status = unzip $name => \$z;
		$data1 = $z;
	}

	# Convert everything to UTF so that decode_entities can handle it
	$data1 = decode("iso-8859-1",$data1);
	# We may not need to use decode_entities.  Can be removed easily.
	$data1 = decode_entities($data1);
	# Sometimes .jpz files have lots and lots of spaces.  Shrink them down.
	$data1 =~ s/\s+/ /g;

	my $xs = XML::Simple->new();
	my $xml = $xs->XMLin($data1);

	###
	#print Dumper($xml);
	#exit(0);
	###
	
	my %puzzle_data;

	$puzzle_data{'title'} = $xml->{'rectangular-puzzle'}->{'metadata'}->{'title'};
	$puzzle_data{'author'} = $xml->{'rectangular-puzzle'}->{'metadata'}->{'creator'};
	$puzzle_data{'copyright'} = $xml->{'rectangular-puzzle'}->{'metadata'}->{'copyright'};
	# Get rid of the annoying copyright symbol
	my $cpr = chr(65533);
	my $cpr2 = chr(169);
	$puzzle_data{'copyright'} =~ s/[$cpr]/(c)/;
	my $x = $puzzle_data{'copyright'};
	$x = substr($x,0,1);
	$puzzle_data{'notes'} = $xml->{'rectangular-puzzle'}->{'instructions'};

	my $gridinfo = $xml->{'rectangular-puzzle'}->{'crossword'}->{'grid'};
	my $w = $gridinfo->{'width'};
	my $h = $gridinfo->{'height'};

	my $cells = $gridinfo->{'cell'};
	foreach my $cell (@$cells) {
		my $r = $cell->{'y'} - 1;
		my $c = $cell->{'x'} - 1;
		if ($cell->{'type'}) {
			if ($cell->{'type'} eq 'block') {
				$puzzle_data{'puzzle'}[$r][$c] = '#';
				$puzzle_data{'solution'}[$r][$c] = '#';
			}
		}
		else {
			# Check for circles
			if ($cell->{'background-shape'}) {
				if ($cell->{'background-shape'} eq 'circle') {
					my %circle_hash;
					$circle_hash{'shapebg'} = 'circle';
					$puzzle_data{'puzzle'}[$r][$c]{'style'} = {%circle_hash};
					# Check for number
					if ($cell->{'number'}) {
						$puzzle_data{'puzzle'}[$r][$c]{'cell'} = int($cell->{'number'});
					}
					else {
						$puzzle_data{'puzzle'}[$r][$c]{'cell'} = 0;
					}
				}
			}
			else { # No circle
				if ($cell->{'number'}) {
					$puzzle_data{'puzzle'}[$r][$c] = int($cell->{'number'});
				}
				else {$puzzle_data{'puzzle'}[$r][$c] = 0;}
			}
			
			# Check the cell solution
			$puzzle_data{'solution'}[$r][$c] = $cell->{'solution'};
		} # end else
	} # end foreach @$cells

	#foreach (@solution) {
	#	print "$_\n";
	#}
	#exit(0);

	# Now do the clues.
	unless ($xml->{'rectangular-puzzle'}->{'crossword'}->{'clues'}->[0]->{'ordering'} eq 'normal') {
		die "$0: We can't handle screwed-up clue ordering yet.  Sorry.";
	}
	my $clues = $xml->{'rectangular-puzzle'}->{'crossword'}->{'clues'};
	my @across;
	my @down;
	# Assume Across comes before down (scary, I know ... but it gets hard to parse here)
	my $ac = $clues->[0]->{'clue'};
	$puzzle_data{'clues'}->{'Across'} = ();
	foreach my $clue (@$ac) {
		my $n = $clue->{'number'};
		delete $clue->{'number'};
		delete $clue->{'word'};
		for my $k (keys %$clue) {
			# Ugh, sometimes these are bolded or italicized.
			my $clux = $clue->{$k};
			while (ref($clux) eq "HASH") {
				if ($clux->{'b'}) {$clux = $clux->{'b'};}
				elsif ($clux->{'i'}) {$clux = $clux->{'i'};}
				else {die "I don't recognize a clue type here.";}
			}	
			#$across[$n] = $clux;
			push(@{$puzzle_data{'clues'}->{'Across'}},[(int($n),$clux)]);
		}
	}
	
	my $dn = $clues->[1]->{'clue'};
	$puzzle_data{'clues'}->{'Down'} = ();
	foreach my $clue (@$dn) {
		my $n = $clue->{'number'};
		delete $clue->{'number'};
		delete $clue->{'word'};
		for my $k (keys %$clue) {
			# Ugh, sometimes these are bolded or italicized.
			my $clux = $clue->{$k};
			while (ref($clux) eq "HASH") {
				if ($clux->{'b'}) {$clux = $clux->{'b'};}
				elsif ($clux->{'i'}) {$clux = $clux->{'i'};}
				else {die "I don't recognize a clue type here.";}
			}	
			#$down[$n] = $clux;
			push(@{$puzzle_data{'clues'}->{'Down'}},[(int($n),$clux)]);
		}
	}
	
	return %puzzle_data;
}