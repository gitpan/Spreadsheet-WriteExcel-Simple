package Spreadsheet::WriteExcel::Simple;

use strict;
use vars qw/$VERSION/;
$VERSION = 0.01;

use Spreadsheet::WriteExcel 0.31;
use IO::Scalar              1.126;

=head1 NAME

Spreadsheet::WriteExcel::Simple - A simple single-sheet Excel document

=head1 SYNOPSIS

  my $ss = Spreadsheet::WriteExcel::Simple->new;
     $ss->write_bold_row(@headings);
     $ss->write_row(@data);
  print $ss->data;

=head1 DESCRIPTION

This provides an abstraction to the Spreadsheet::WriteExcel module
for easier creation of simple single-sheet Excel documents.

In its most basic form it provides two methods for writing data:
write_row and write_bold_row which write the data supplied to
the next row of the spreadsheet. 

However, you can also use $ss->book and $ss->sheet to get at the
underlying workbook and worksheet from Spreadsheet::WriteExcel if you
wish to manipulate these directly.

=head1 METHODS

=head2 new

  my $ss = Spreadsheet::WriteExcel::Simple->new;

Create a new single-sheet Excel document. You do not need to supply
this a filename or filehandle. The data is store internally, and can
be retrieved later through the 'data' method.

=cut

sub new {
  my $class = shift;
  my $self = bless {}, $class;

  my $fh = shift;
  # Store the workbook in a tied scalar filehandle
  $self->{book} = Spreadsheet::WriteExcel->new(
    IO::Scalar->new_tie(\($self->{content}))
  );
  $self->{bold} = $self->book->addformat();
  $self->{bold}->set_bold;
  $self->{sheet} = $self->book->addworksheet;
  $self;
}

=head2 write_row / write_bold_row

  $ss->write_bold_row(@headings);
  $ss->write_row(@data);

These write the list of data into the next row of the spreadsheet.

Caveat: An internal counter is kept as to which row is being written
to, so if you mix these functions with direct writes of your own,
these functions will continue where they left off, not where you have
written to.

=cut

{ my $row = 0;

  sub write_row {
    my $self = shift;
    my $dataref = shift;
    my @data = map { defined $_ ? $_ : '' } @$dataref;
    my $fmt  = shift || '';
    my $col = 0;
    my $ws = $self->sheet;
       $ws->write($row, $col++, $_, $fmt) foreach @data;
    $row++;
  }

  sub write_bold_row { $_[0]->write_row($_[1], $_[0]->_bold) }
}

=head2 data

  print $ss->data;

This returns the data of the spreadsheet. If you're planning to print this
to a web-browser, be sure to print an 'application/excel' header first.

=cut

sub data {
  my $self = shift;
  $self->book->close;
  return $self->{content};
}

=head2 book / sheet

  my $workbook  = $ss->book;
  my $worksheet = $ss->sheet;

These return the underlying Spreadsheet::WriteExcel objects representing
the workbook and worksheet respectively. If you find yourself doing a
lot of work with these, you probably shouldn't be using this module,
but using Spreadsheet::WriteExcel directly.

=cut

sub book  { $_[0]->{book} }
sub sheet { $_[0]->{sheet} }

sub _bold { $_[0]->{bold} }

=head1 BUGS

This can't yet handle dates.

=head1 AUTHOR

Tony Bowden, E<lt>tony@tmtm.comE<gt>.

=head1 SEE ALSO

L<Spreadsheet::WriteExcel>. John McNamara has done a great job with this.

=head1 COPYRIGHT

Copyright (C) 2001 Tony Bowden. All rights reserved.

This module is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=cut

1;

