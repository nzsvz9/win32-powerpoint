package Win32::PowerPoint;

use strict;
use warnings;
use Carp;

our $VERSION = '0.20';

use File::Spec;
use File::Basename;
use Win32::OLE;
use Win32::OLE::Variant;
use Win32::PowerPoint::Constants;
use Win32::PowerPoint::Utils qw(
  RGB
  canonical_alignment
  canonical_pattern
  canonical_datetime
  convert_cygwin_path
  _defined_or
);

sub new {
  my $class = shift;
  my $self  = bless {
    c            => Win32::PowerPoint::Constants->new,
    was_invoked  => 0,
    application  => undef,
    presentation => undef,
    slide        => undef,
  }, $class;

  $self->connect_or_invoke;

  return $self;
}

sub c { shift->{c} }

##### application #####

sub application { shift->{application} }

sub connect_or_invoke {
  my $self = shift;

  $self->{application} = Win32::OLE->GetActiveObject('PowerPoint.Application');

  unless (defined $self->{application}) {
    $self->{application} = Win32::OLE->new('PowerPoint.Application')
      or die Win32::OLE->LastError;
    $self->{was_invoked} = 1;
  }
}

sub quit {
  my $self = shift;

  return unless $self->application;

  $self->application->Quit;
  $self->{application} = undef;
}

##### presentation #####

sub new_presentation {
  my $self = shift;

  return unless $self->{application};

  my %options = ( @_ == 1 and ref $_[0] eq 'HASH' ) ? %{ $_[0] } : @_;

  $self->{slide} = undef;

  $self->{presentation} = $self->application->Presentations->Add
    or die Win32::OLE->LastError;

  $self->_apply_background(
    $self->presentation->SlideMaster->Background->Fill,
    %options
  );
}

sub presentation {
  my $self = shift;

  return unless $self->{application};

  $self->{presentation} ||= $self->application->ActivePresentation
    or die Win32::OLE->LastError;
}

sub _apply_background {
  my ($self, $target, %options) = @_;

  my $forecolor = _defined_or(
    $options{background_forecolor},
    $options{masterbkgforecolor}
  );
  if ( defined $forecolor ) {
    $target->ForeColor->{RGB} = RGB($forecolor);
    $self->slide->{FollowMasterBackground} = $self->c->msoFalse if $options{slide};
  }

  my $backcolor = _defined_or(
    $options{background_backcolor},
    $options{masterbkgbackcolor}
  );
  if ( defined $backcolor ) {
    $target->BackColor->{RGB} = RGB($backcolor);
    $self->slide->{FollowMasterBackground} = $self->c->msoFalse if $options{slide};
  }

  if ( defined $options{pattern} ) {
    if ( $options{pattern} =~ /\D/ ) {
      my $method = canonical_pattern($options{pattern});
      $options{pattern} = $self->c->$method;
    }
    $target->Patterned( $options{pattern} );
  }
}

sub save_presentation {
  my ($self, $file) = @_;

  return unless $self->presentation;
  return unless defined $file;

  my $absfile   = File::Spec->rel2abs($file);
  my $directory = dirname( $file );
  unless (-d $directory) {
    require File::Path;
    File::Path::mkpath($directory);
  }

  $self->presentation->SaveAs( convert_cygwin_path( $absfile ) );
}

sub close_presentation {
  my $self = shift;

  return unless $self->presentation;

  $self->presentation->Close;
  $self->{presentation} = undef;
}

sub set_master_footer {
  my $self = shift;

  return unless $self->presentation;
  my $master_footers = $self->presentation->SlideMaster;
  $self->_set_footer($master_footers, @_);
}

sub _set_footer {
  my ($self, $slide, @args) = @_;

  my $target = $slide->HeadersFooters;

  my %options = ( @args == 1 and ref $args[0] eq 'HASH' ) ? %{ $args[0] } : @args;

  if ( defined $options{visible} ) {
    $target->Footer->{Visible} = $options{visible} ? $self->c->msoTrue : $self->c->msoFalse;
  }

  if ( defined $options{text} ) {
    $target->Footer->{Text} = $options{text};
  }

  if ( defined $options{slide_number} ) {
    $target->SlideNumber->{Visible} = $options{slide_number} ? $self->c->msoTrue : $self->c->msoFalse;
  }

  if ( defined $options{datetime} ) {
    $target->DateAndTime->{Visible} = $options{datetime} ? $self->c->msoTrue : $self->c->msoFalse;
  }

  if ( defined $options{datetime_format} ) {
    if ( !$options{datetime_format} ) {
      $target->DateAndTime->{UseFormat} = $self->c->msoFalse;
    }
    else {
      if ( $options{datetime_format} =~ /\D/ ) {
        my $format = canonical_datetime($options{datetime_format});
        $options{datetime_format} = $self->c->$format;
      }
      $target->DateAndTime->{UseFormat} = $self->c->msoTrue;
      $target->DateAndTime->{Format}    = $options{datetime_format};
    }
  }
}

##### slide #####

sub slide {
  my ($self, $id) = @_;
  if ($id) {
    $self->{slide} = $self->presentation->Slides->Item($id)
      or die Win32::OLE->LastError;
  }
  $self->{slide};
}

sub new_slide {
  my $self = shift;

  my %options = ( @_ == 1 and ref $_[0] eq 'HASH' ) ? %{ $_[0] } : @_;

  $self->{slide} = $self->presentation->Slides->Add(
    $self->presentation->Slides->Count + 1,
    $self->c->LayoutBlank
  ) or die Win32::OLE->LastError;
  $self->{last} = undef;

  $self->_apply_background(
    $self->slide->Background->Fill,
    %options,
    slide => 1,
  );
}


#####    ##    ####  ######         ####  ###### ##### #    # #####
#    #  #  #  #    # #             #      #        #   #    # #    #
#    # #    # #      #####          ####  #####    #   #    # #    #
#####  ###### #  ### #                  # #        #   #    # #####
#      #    # #    # #             #    # #        #   #    # #
#      #    #  ####  ###### ######  ####  ######   #    ####  #

# Added 03Sep2022 by Thomas Catsburg
#
# Special thanks to Perl Monks Bod and kcott and an Anonymous Monk for their insights into page setup

sub page_setup
{
  # Get the self and options
  my ($self, $options) = @_;

  # Clear options unless the given option set is a hash
  $options = {} unless ref $options eq 'HASH';

  # PPT.PageSetup.SlideSize = ppSlideSizeA4Paper
  # Const ppSlideSizeA4Paper = 3

  # Microsoft SlideSize in Constants.pm
  #
  #  Name Value Description
  #  ppSlideSize35MM  4 35MM
  #  ppSlideSizeA3Paper 9 A3 Paper
  #  ppSlideSizeA4Paper 3 A4 Paper
  #  ppSlideSizeB4ISOPaper  10  B4 ISO Paper
  #  ppSlideSizeB4JISPaper  12  B4 JIS Paper
  #  ppSlideSizeB5ISOPaper  11  B5 ISO Paper
  #  ppSlideSizeB5JISPaper  13  B5 JIS Paper
  #  ppSlideSizeBanner  6 Banner
  #  ppSlideSizeCustom  7 Custom
  #  ppSlideSizeHagakiCard  14  Hagaki Card
  #  ppSlideSizeLedgerPaper 8 Ledger Paper
  #  ppSlideSizeLetterPaper 2 Letter Paper
  #  ppSlideSizeOnScreen  1 On Screen
  #  ppSlideSizeOverhead  5 Overhead

  # If given SlideSize
  if ( defined $options->{SlideSize} )
  {
    # Set the slide size
    $self->presentation->PageSetup->{SlideSize} = $options->{SlideSize};
  }

  # If given slide width
  if ( defined $options->{SlideWidth} )
  {
    # Set the slide width
    $self->presentation->PageSetup->{SlideWidth} = $options->{SlideWidth};
  }

  # If given the slide height
  if ( defined $options->{SlideHeight} )
  {
    # Set the slide height
    $self->presentation->PageSetup->{SlideHeight} = $options->{SlideHeight};
  }

  # Return
  return;
}



sub set_footer {
  my $self = shift;

  return unless $self->slide;
  $self->_set_footer($self->slide, @_);
}

sub add_text {
  my ($self, $text, $options) = @_;

  return unless $self->slide;
  return unless defined $text;

  $options = {} unless ref $options eq 'HASH';

  $text =~ s/\n/\r/gs;

  my ($left, $top, $width, $height);
  if (my $last = $self->{last}) {
    $left   = _defined_or($options->{left},   $last->Left);
    $top    = _defined_or($options->{top},    $last->Top + $last->Height + 20);
    $width  = _defined_or($options->{width},  $last->Width);
    $height = _defined_or($options->{height}, $last->Height);
  }
  else {
    $left   = _defined_or($options->{left},   30);
    $top    = _defined_or($options->{top},    30);
    $width  = _defined_or($options->{width},  600);
    $height = _defined_or($options->{height}, 200);
  }

  my $new_textbox = $self->slide->Shapes->AddTextbox(
    $self->c->TextOrientationHorizontal,
    $left, $top, $width, $height
  );

  my $frame = $new_textbox->TextFrame;
  my $range = $frame->TextRange;

  $frame->{WordWrap} = $self->c->True;
  $range->ParagraphFormat->{FarEastLineBreakControl} = $self->c->True;
  $range->{Text} = $text;

  $self->decorate_range( $range, $options );

  $frame->{AutoSize} = $self->c->AutoSizeNone;
  $frame->{AutoSize} = $self->c->AutoSizeShapeToFitText;

  $self->{last} = $new_textbox;

  return $new_textbox;
}


  ##   #####  #####         #      ### #    # ######
 #  #  #    # #    #        #       #  ##   # #
#    # #    # #    #        #       #  # #  # #####
###### #    # #    #        #       #  #  # # #
#    # #    # #    #        #       #  #   ## #
#    # #####  #####  ###### ###### ### #    # ######

# Added 03Sep2022 by Thomas Catsburg
#
# Special thanks to Perl Monk haukex for insight on how to correctly call Shapes->AddLine
#

sub add_line
{
  # Get the self and options
  my ($self, $options) = @_;

  # Return if self is not a slide
  return unless $self->slide;

  # Clear options unless the given option set is a hash
  $options = {} unless ref $options eq 'HASH';

  # Create the line using the given 4 points
  my $new_line = $self->slide->Shapes->AddLine($options->{x1}, $options->{y1}, $options->{x2}, $options->{y2});

  # Set the line color if given
  $new_line->{Line}{ForeColor}{RGB}=RGB($options->{forecolor}) if(defined $options->{forecolor});

  # Set the line weight if given
  $new_line->{Line}{Weight}=$options->{weight} if(defined $options->{weight});

  # Set the line pattern if given
  $new_line->{Line}{Pattern}=$options->{pattern} if(defined $options->{pattern});

  # Return the line handle
  return $new_line;
}


  ##   #####  #####         #####   ####  #       #   # #      ### #    # ######
 #  #  #    # #    #        #    # #    # #        # #  #       #  ##   # #
#    # #    # #    #        #    # #    # #         #   #       #  # #  # #####
###### #    # #    #        #####  #    # #         #   #       #  #  # # #
#    # #    # #    #        #      #    # #         #   #       #  #   ## #
#    # #####  #####  ###### #       ####  ######    #   ###### ### #    # ######

# Added 14Jun2023 by Thomas Catsburg

sub add_polyline
{
  # Get the self and options
  my ($self, $options) = @_;

  # Return if self is not a slide
  return unless $self->slide;

  # Clear options unless the given option set is a hash
  $options = {} unless ref $options eq 'HASH';

  my $points=$options->{points};

  # Get the points
  my @points=@{ $options->{points} };

  # Add the first 2 points onto the end of the list to close the loop
  push @points, $points[0];
  push @points, $points[1];

  # Choose an array size
  my $asize=int(($#points+1)/2);

  # Create the Win32::OLE::Variant - Special thanks to Perl Monk Corion for help on how to dimension the pointlist array
  my $pointlist = Win32::OLE::Variant->new(VT_ARRAY | VT_R4 , $asize, 2);

  # Convert points to Win32::OLE::Variant
  for my $index (0 .. int(($#points-1)/2))
  {
    # Add x
    $pointlist->Put( $index, 0, shift(@points) );

    # Add y
    $pointlist->Put( $index, 1, shift(@points) );
  }

  # Create the poly line using the variant point list
  my $new_poly=$self->slide->Shapes->AddPolyline($pointlist);

  # Set the line color if given
  $new_poly->{Line}{ForeColor}{RGB}=RGB($options->{forecolor}) if(defined $options->{forecolor});

  # Set the line weight if given
  $new_poly->{Line}{Weight}=$options->{weight} if(defined $options->{weight});

  # Set the line pattern if given
  $new_poly->{Line}{Pattern}=$options->{pattern} if(defined $options->{pattern});

  # If the fill color is given
  if(defined $options->{fillcolor})
  {
    # Set the fill color
    $new_poly->{Fill}{ForeColor}{RGB}=RGB($options->{fillcolor});
  }
  else
  {
    # Kludge - set transparancy to 100% if no fill color given
    $new_poly->{Fill}{Transparency}=1;
  }

  # Return the poly line handle
  return $new_poly;
}


  ##   #####  #####          ####  #    # #####  #    # ######
 #  #  #    # #    #        #    # #    # #    # #    # #
#    # #    # #    #        #      #    # #    # #    # #####
###### #    # #    #        #      #    # #####  #    # #
#    # #    # #    #        #    # #    # #   #   #  #  #
#    # #####  #####  ######  ####   ####  #    #   ##   ######

# Added 14Jun2023 by Thomas Catsburg

sub add_curve
{
  # Get the self and options
  my ($self, $options) = @_;

  # Return if self is not a slide
  return unless $self->slide;

  # Clear options unless the given option set is a hash
  $options = {} unless ref $options eq 'HASH';

  my $points=$options->{points};

  # Get the points
  my @points=@{ $options->{points} };

  # Choose an array size
  my $asize=int(($#points+1)/2);

  # Create the Win32::OLE::Variant - Special thanks to Perl Monk Corion for help on how to dimension the pointlist array
  my $pointlist = Win32::OLE::Variant->new(VT_ARRAY | VT_R4 , $asize, 2);

  # Convert points to Win32::OLE::Variant
  for my $index (0 .. int(($#points-1)/2))
  {
    # Add x
    $pointlist->Put( $index, 0, shift(@points) );

    # Add y
    $pointlist->Put( $index, 1, shift(@points) );
  }

  # Create the curve line using the variant point list
  my $new_curve=$self->slide->Shapes->AddCurve($pointlist);

  # Set the line color if given
  $new_curve->{Line}{ForeColor}{RGB}=RGB($options->{forecolor}) if(defined $options->{forecolor});

  # Set the line weight if given
  $new_curve->{Line}{Weight}=$options->{weight} if(defined $options->{weight});

  # Set the line pattern if given
  $new_curve->{Line}{Pattern}=$options->{pattern} if(defined $options->{pattern});

  # If the fill color is given
  if(defined $options->{fillcolor})
  {
    # Set the fill color
    $new_curve->{Fill}{ForeColor}{RGB}=RGB($options->{fillcolor});
  }
  else
  {
    # Kludge - set transparancy to 100% if no fill color given
    $new_curve->{Fill}{Transparency}=1;
  }

  # Return the curve line handle
  return $new_curve;
}


  ##   #####  #####          ####  #    #   ##   #####  ######
 #  #  #    # #    #        #      #    #  #  #  #    # #
#    # #    # #    #         ####  ###### #    # #    # #####
###### #    # #    #             # #    # ###### #####  #
#    # #    # #    #        #    # #    # #    # #      #
#    # #####  #####  ######  ####  #    # #    # #      ######

# Added 03Sep2022 by Thomas Catsburg
#
# Special thanks to Perl Monk haukex for insight on how to correctly call Shapes->AddShape
#

sub add_shape
{
  # Get self shape and options
  my ($self, $shape, $options) = @_;

  # Return if self is not a slide
  return unless $self->slide;

  # Clear options unless the given option set is a hash
  $options = {} unless ref $options eq 'HASH';

  # Create the shape using the given shape and options
  my $new_shape = $self->slide->Shapes->AddShape($self->c->$shape, $options->{left}, $options->{top}, $options->{width}, $options->{height});

  # Set the line color if given
  $new_shape->{Line}{ForeColor}{RGB}=RGB($options->{bordercolor}) if(defined $options->{bordercolor});

  # Set the line weight if given
  $new_shape->{Line}{Weight}=$options->{weight} if(defined $options->{weight});

  # If the fill color is given
  if(defined $options->{fillcolor})
  {
    # Set the fill color
    $new_shape->{Fill}{ForeColor}{RGB}=RGB($options->{fillcolor});
  }
  else
  {
    # Kludge - set transparancy to 100% if no fill color given
    $new_shape->{Fill}{Transparency}=1;
  }

  # Set the shape rotation if given
  $new_shape->{Rotation}=$options->{rotation} if(defined $options->{rotation});

  # If shape is a msoShapeArc
  if($shape =~ /msoShapeArc/i)
  {
    # Set the start and extent angles of an arc
    $new_shape->{Adjustments}{1}=$options->{start};
    $new_shape->{Adjustments}{2}=$options->{extent};
  }

  # Return the shape handle
  return $new_shape;
}


  ##   #####  #####         #####  ###  ####  ##### #    # #####  ######
 #  #  #    # #    #        #    #  #  #    #   #   #    # #    # #
#    # #    # #    #        #    #  #  #        #   #    # #    # #####
###### #    # #    #        #####   #  #        #   #    # #####  #
#    # #    # #    #        #       #  #    #   #   #    # #   #  #
#    # #####  #####  ###### #      ###  ####    #    ####  #    # ######

sub add_picture
{
  # Get the self image file and options
  my ($self, $file, $options) = @_;

  # Return if self is not a slide
  return unless $self->slide;

  # Return if file is not a file
  return unless defined $file and -f $file;

  # Clear options unless the given option set is a hash
  $options = {} unless ref $options eq 'HASH';

  my ($left, $top);
  if (my $last = $self->{last}) {
    $left   = _defined_or($options->{left}, $last->Left);
    $top    = _defined_or($options->{top},  $last->Top + $last->Height + 20);
  }
  else {
    $left   = _defined_or($options->{left}, 30);
    $top    = _defined_or($options->{top},  30);
  }

  my $new_picture = $self->slide->Shapes->AddPicture(
    convert_cygwin_path( $file ),
    ( $options->{link}
      ? ( $self->c->msoTrue,  $self->c->msoFalse )
      : ( $self->c->msoFalse, $self->c->msoTrue )
    ),
    $left, $top, $options->{width}, $options->{height}
  );

  $self->{last} = $new_picture;

  return $new_picture;
}

sub insert_before {
  my ($self, $text, $options) = @_;

  return unless $self->slide;
  return unless defined $text;

  $options = {} unless ref $options eq 'HASH';

  $text =~ s/\n/\r/gs;

  my $num_of_boxes = $self->slide->Shapes->Count;
  my $last  = $num_of_boxes ? $self->slide->Shapes($num_of_boxes) : undef;
  my $range = $self->slide->Shapes($num_of_boxes)->TextFrame->TextRange;

  my $selection = $range->InsertBefore($text);

  $self->decorate_range( $selection, $options );

  return $selection;
}

sub insert_after {
  my ($self, $text, $options) = @_;

  return unless $self->slide;
  return unless defined $text;

  $options = {} unless ref $options eq 'HASH';

  $text =~ s/\n/\r/gs;

  my $num_of_boxes = $self->slide->Shapes->Count;
  my $last  = $num_of_boxes ? $self->slide->Shapes($num_of_boxes) : undef;
  my $range = $self->{slide}->Shapes($num_of_boxes)->TextFrame->TextRange;

  my $selection = $range->InsertAfter($text);

  $self->decorate_range( $selection, $options );

  return $selection;
}

sub decorate_range {
  my ($self, $range, $options) = @_;

  return unless defined $range;

  $options = {} unless ref $options eq 'HASH';

  my ($true, $false) = ($self->c->True, $self->c->False);

  $range->Font->{Bold}        = $options->{bold}        ? $true : $false;
  $range->Font->{Italic}      = $options->{italic}      ? $true : $false;
  $range->Font->{Underline}   = $options->{underline}   ? $true : $false;
  $range->Font->{Shadow}      = $options->{shadow}      ? $true : $false;
  $range->Font->{Subscript}   = $options->{subscript}   ? $true : $false;
  $range->Font->{Superscript} = $options->{superscript} ? $true : $false;
  $range->Font->{Size}        = $options->{size}       if $options->{size};
  $range->Font->{Name}        = $options->{name}       if $options->{name};
  $range->Font->{Name}        = $options->{font}       if $options->{font};
  $range->Font->Color->{RGB}  = RGB($options->{color}) if $options->{color};

  my $align = $options->{alignment} || $options->{align} || 'left';
  if ( $align =~ /\D/ ) {
    my $method = canonical_alignment( $align );
    $align = $self->c->$method;
  }
  $range->ParagraphFormat->{Alignment} = $align;

  $range->ActionSettings(
    $self->c->MouseClick
  )->Hyperlink->{Address} = $options->{link} if $options->{link};
}

sub DESTROY {
  my $self = shift;

  $self->quit if $self->{was_invoked};
}

1;
__END__

=head1 NAME

Win32::PowerPoint - Create and Edit PowerPoint presentations

=head1 SYNOPSIS

    use Win32::PowerPoint;

    # invoke (or connect to) PowerPoint
    my $pp = Win32::PowerPoint->new;

    # set presentation-wide information
    $pp->new_presentation(
      background_forecolor => [255,255,255],
      background_backcolor => 'RGB(0, 0, 0)',
      pattern => 'Shingle',
    );

    # and master footer if you prefer (optional)
    $pp->set_master_footer(
      visible         => 1,
      text            => 'My Slides',
      slide_number    => 1,
      datetime        => 1,
      datetime_format => 'MMMMyy',
    );

    (load and parse your slide)

    # do whatever you want to do for each of your slides
    foreach my $slide (@slides) {
      $pp->new_slide;

      $pp->add_text($slide->title, { size => 40, bold => 1 });
      $pp->add_text($slide->body);
      $pp->add_text($slide->link,  { link => $slide->link });

      # you may add pictures
      $pp->add_picture($file, { left => 10, top => 10 });
    }

    $pp->save_presentation('slide.ppt');

    $pp->close_presentation;

    # PowerPoint closes automatically

=head1 DESCRIPTION

Win32::PowerPoint helps you to create or edit a PowerPoint presentation. You can add text/images incrementally to your slides.

=head1 METHODS

=head2 new

Invokes (or connects to) PowerPoint.

    my $pp=Win32::PowerPoint->new();

=head2 connect_or_invoke

Explicitly connects to (or invoke) PowerPoint.

=head2 quit

Explicitly disconnects from PowerPoint, and closes it if this module invoked it.

=head2 new_presentation (options)

Creates a new (probably blank) presentation. Options are:

=over 4

=item background_forecolor, background_backcolor

You can specify background colors of the slides with an array ref of RGB
components ([255, 255, 255] for white) or formatted string ('255, 0, 0'
for red). You can use '(0, 255, 255)' or 'RGB(0, 255, 255)' format for
clarity. These colors are applied to all the slides you'll add, unless
you specify other colors for the slides explicitly.

You can use 'masterbkgforecolor' and 'masterbkgbackcolor' as aliases.

Extension: Use the routine convert2RGBvalues to take a color name or
RGB hex color and convert to a comma seperated list of decimal color
values so black becomes '0, 0, 0'

=item pattern

You also can specify default background pattern for the slides.
See L<Win32::PowerPoint::Constants> (or MSDN or PowerPoint's help) for
supported pattern names. You can omit 'msoPattern' part and the names
are case-sensitive.

=back

=head2 page_setup

Set the page dimensions

  Standard 4:3 PowerPoint page is 10in x 7.5in or 720pt x 540pt

  $pp->page_setup( { SlideWidth  => 720,
                     SlideHeight => 540 });

=head2 save_presentation (path)

Saves the presentation to where you specified. Accepts relative path.
You might want to save it as .pps (slideshow) file to make it easy to
show slides (it just starts full screen slideshow with a doubleclick).

=head2 close_presentation

Explicitly closes the presentation.

=head2 new_slide (options)

Adds a new (blank) slide to the presentation. Options are:

=over 4

=item background_forecolor, background_backcolor

You can set colors just for the slide with these options.
You can use 'bkgforecolor' and 'bkgbackcolor' as aliases.

=item pattern

You also can set background pattern just for the slide.

=back

=head2 add_text (text, options)

Adds (formatted) text to the slide. Options are:

=over 4

=item left, top, width, height

of the Textbox.

=back

See 'decorate_range' for other options.

=head2 add_line

Draw a line with coordinate pairs with defined weight (in points) and color

  $pp->add_line({'x1'        => $x1,
                 'y1'        => $y1,
                 'x2'        => $x2,
                 'y2'        => $y2,
                 'weight'    => 1,
                 'forecolor' => &convert2RGBvalues($color) });

=head2 add_polyline

Draw a line with multiple coordinate pairs, closing the line by repeating the 
initial coordinate pair to the end of the list of coordinate pairs, with 
defined weight (in points) and color and fillcolor.

This method is used to draw shapes using point sets, with three points making a
triangle, 4 a rectanlge and so on.

  $pp->add_polyline({'points'    => [@points[0 .. $#points]],
                     'fillcolor' => &convert2RGBvalues($fill),
                     'weight'    => 1,
                     'forecolor' => &convert2RGBvalues($color) });

=head2 add_curve

Draw a smooth curve spline given a set of at least 3 coordinate pairs with
defined weight (in points) and color.

  $pp->add_curve({'points'    => [@curve[0 .. $#curve]],
                  'weight'    => 1,
                  'forecolor' => &convert2RGBvalues($color) } );

=head2 add_shape

Draw PowerPoint pre-defined shape as from the PowerPoint Insert >> Shape menu.

  The shapes are given in Constants.pm as PowerPoint shape names msoShape
  See L<Win32::PowerPoint::Constants> for msoAutoShapeType

=head2 add_picture (file, options)

Adds file to the slide. Options are:

=over 4

=item left, top, width, height

of the picture. width and height are optional.

=item link

If set to true, the picture will be linked, otherwise, embedded.

=back

=head2 insert_before (text, options)

=head2 insert_after (text, options)

Prepends/Appends text to the current Textbox. See 'decorate_range' for options.

=head2 set_footer, set_master_footer (options)

Arranges (master) footer. Options are:

=over 4

=item visible

If set to true, the footer(s) will be shown, and vice versa.

=item text

Specifies the text part of the footer(s)

=item slide_number

If set to true, slide number(s) will be shown, and vice versa.

=item datetime

If set to true, the date time part of the footer(s) will be shown, and vice versa.

=item datetime_format

Specifies the date time format of the footer(s) if you specify one of the registered ppDateTimeFormat name (see L<Win32::PowerPoint::Constants> or MSDN for details). If set to false, no format will be used.

=back

=head2 decorate_range (range, options)

Decorates text of the range. Options are:

=over 4

=item bold, italic, underline, shadow, subscript, superscript

Boolean.

=item size

Integer.

=item color

See above for the convention.

=item font

Font name of the text. You can use 'name' as an alias.

=item alignment

One of the 'left' (default), 'center', 'right', 'justify', 'distribute'.

You can use 'align' as an alias.

=item link

hyperlink address of the Text.

=back

(This method is mainly for the internal use).

=head1 IF YOU WANT TO GO INTO DETAIL

This module uses L<Win32::OLE> internally. You can fully control PowerPoint through the following accessors. See L<Win32::OLE> and other appropriate documents like intermediate books on PowerPoint and Visual Basic for details (after all, this module is just a thin wrapper of them). If you're still using old PowerPoint (2003 and older), try C<Record New Macro> (from the C<Tools> menu, then, C<Macro>, and voila) and do what you want, and see what's recorded (from the C<Tools> menu, then C<Macro>, and C<Macro...> submenu. You'll see Visual Basic Editor screen).

=head2 application

returns an Application object.

    print $pp->application->Name;

=head2 presentation

returns a current Presentation object (maybe ActivePresentation but that's not assured).

    $pp->save_presentation('sample.ppt') unless $pp->presentation->Saved;

    while (my $last = $pp->presentation->Slides->Count) {
      $pp->presentation->Slides($last)->Delete;
    }

=head2 slide

returns a current Slide object.

    $pp->slide->Export(".\\slide_01.jpg",'jpg');

    $pp->slide->Shapes(1)->TextFrame->TextRange
       ->Characters(1, 5)->Font->{Bold} = $pp->c->True;

As of 0.10, you can pass an index number to get an arbitrary Slide object.

=head2 c

returns Win32::PowerPoint::Constants object.

=head1 CAVEATS FOR CYGWIN USERS

This module itself seems to work under the Cygwin environment. However, MS PowerPoint expects paths to be Windows-ish, namely without /cygdrive/. So, when you load or save a presentation, or import some materials with OLE (native) methods, you usually need to convert them by yourself. As of 0.08, Win32::PowerPoint::Utils has a C<convert_cygwin_path> function for this. Win32::PowerPoint methods use this function internally, so you don't need to convert paths explicitly.

=head1 AUTHOR

Kenichi Ishigaki, E<lt>ishigaki@cpan.orgE<gt>

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2006 - by Kenichi Ishigaki
              2023 - by Thomas Catsburg

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=cut
