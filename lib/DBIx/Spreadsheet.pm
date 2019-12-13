#!perl
package DBIx::Spreadsheet;
use strict;
use DBI;
use Getopt::Long;
use Spreadsheet::Read;
use Text::CleanFragment;

use Moo 2;

use Filter::signatures;
no warnings 'experimental::signatures';
use feature 'signatures';

our $VERSION = '0.01';

=head1 NAME

DBIx::Spreadsheet - Query a spreadsheet with SQL

=head1 SYNOPSIS

  my $sheet = DBIx::Spreadsheet->new( file => 'workbook.xlsx' );
  my $dbh = $sheet->dbh;

  my @rows = $dbh->selectall_arrayref(<<'SQL');
      select *
        from sheet_1
       where foo = 'bar'
  SQL

This module reads a workbook and makes the contained spreadsheets available
as tables. It assumes that the first row of a spreadsheet are the column
names. Empty column names will be replaced by C<col_$number>. The column
names will be sanitized by L<Text::CleanFragment> so they are conveniently
usable.

=head1 METHODS

=head2 C<< DBIx::Spreadsheet->new >>

  my $wb = DBIx::Spreadsheet->new(
      file => 'workboook.ods',
  );

=head3 Options

=over 4

=item *

B<file> - name of the workbook file. The file will be read using L<Spreadsheet::Read>
using the options in C<spreadsheet_options>.

=cut

has 'file' => (
    is => 'ro',
);

=item *

B<spreadsheet> - a premade L<Spreadsheet::Read> object

=cut

has 'spreadsheet' => (
    is => 'lazy',
    default => \&_read_file,
);

=item *

B<spreadsheet_options> - options for the L<Spreadsheet::Read> object

=back

=cut

has 'spreadsheet_options' => (
    is => 'lazy',
    default => sub { {
        dtfmt => 'yyyy-mm-dd',
    } },
);

=head2 C<< ->dbh >>

  my $dbh = $wb->dbh;

Returns the database handle to access the sheets.

=cut

has 'dbh' => (
    is => 'lazy',
    default => sub( $self ) { $self->_import_data; $self->{dbh} },
);

=head2 C<< ->tables >>

  my $tables = $wb->tables;

Arrayref containing the names of the tables. These are usually the names
of the sheets.

=cut

has 'tables' => (
    is => 'lazy',
    default => sub( $self ) { $self->_import_data; $self->{tables} },
);

sub _read_file( $self ) {
    Spreadsheet::Read->new( $self->file, %{ $self->spreadsheet_options } );
};

our $table_000;

our %charmap = (
    '+' => 'plus',
    '%' => 'perc',
);

sub gen_colnames( $self, @colnames ) {
    my %seen;
    my $i = 1;
    return map { qq{"$_"}}
           map { clean_fragment( $_ ) }
           map { s/([-.])/_/g; $_ }                # replace . and - to _
           map { s/([%+])/_$charmap{ $1 }_/g; $_ } # replace + and % with names
           map { $i++; my $name = $_ eq '' ? sprintf "col_%d", $i : $seen{ $_ } ? "${_}_1" : $_; $seen{$name}++; $name }
           map { !defined($_) ? "" : $_ }
           @colnames;
}

# The nasty fixup until I clean up Spreadsheet::ReadSXC to allow for the raw
# values
sub nasty_cell_fixup( $self, $value ) {
    return $value if ! defined $value;
# use Data::Dumper; $Data::Dumper::Useqq = 1;

    # Fix up German locale formatted numbers, as that's what I have
   if( $value =~ /^([+-]?)([0-9\.]+(,\d+))?(\s*\x{20ac}|€)?$/ ) {
        # Fix up formatted number
        $value =~ s![^\d\.\,+-]!!g;
        $value =~ s!\.!!g;
        $value =~ s!,!.!g;

    # Fix up  German locale formatted dates, as that's what I have
    } elsif( $value =~ /^([0123]?\d)\.([01]\d)\.(\d\d)$/ ) {
        $value = "20$3-$2-$1";

    # Fix up  German locale formatted dates, as that's what I have
    } elsif( $value =~ /^([0123]?\d)\.([01]\d)\.(20\d\d)$/ ) {
        $value = "$3-$2-$1";
    }
    return $value
}

sub import_data( $self, $book ) {
    my $dbh = DBI->connect('dbi:SQLite:dbname=:memory:',undef,undef,{AutoCommit => 1, RaiseError => 1,PrintError => 0});
    $dbh->sqlite_create_module(perl => "DBD::SQLite::VirtualTable::PerlData");

    my @tables;
    my $i = 0;
    for my $table_name ($book->sheets) {
        my $sheet = $book->sheet( $table_name );
        my $tablevar = sprintf 'table_%03d', $i++;
        #warn sprintf "%s: %d, %d", $table_name, $sheet->maxcol, $sheet->maxrow;
        #use Data::Dumper;
        #warn Dumper [$sheet->cellrow(2)];
        #warn Dumper [$sheet->row(2)];
        my $data = [map { [
                            map { $self->nasty_cell_fixup( $_ ) } $sheet->cellrow($_)
                    ] } 1..$sheet->maxrow ];
        #my $data = [$sheet->rows($_)];
        my $colnames = shift @{$data};

        my $sql_name = clean_fragment $table_name;

        # Fix up duplicate columns, empty column names
        $colnames = join ",", $self->gen_colnames( @$colnames );
        {;
            #no strict 'refs';
            # Later, find the first non-empty row, instead of blindly taking the first row
            #${main::}{$tablevar} = \$data;
        };
        local $table_000 = $data;
        $tablevar = __PACKAGE__ . '::table_000';
        my $sql = qq(CREATE VIRTUAL TABLE temp."$sql_name" USING perl($colnames, arrayrefs="$tablevar"););
        $dbh->do($sql);
        push @tables, $sql_name;
    };

    return $dbh, \@tables;
}

sub _import_data( $self ) {
    my( $dbh, $tables ) = $self->import_data( $self->spreadsheet );
    $self->{dbh} = $dbh;
    $self->{tables} = $tables;
}

1;

=head1 SUPPORTED FILE TYPES

This module supports the same file types as L<Spreadsheet::Read>. The following
modules need to be installed to read the various file types:

=over 4

=item *

L<Text::CSV_XS> - CSV files

=item *

L<Spreadsheet::ParseXLS> - Excel XLS files

=item *

L<Spreadsheet::ParseXLSX> - Excel XLSX files

=item *

L<Spreadsheet::ParseSXC> - Staroffice / Libre Office SXC or ODS files

=back

=head1 TO DO

=over 4

=item *

Create DBD so direct usage with L<DBI> becomes possible

  my $dbh = DBI->connect('dbi:Spreadsheet:filename=workbook.xlsx,start_row=2');

DBIx::Spreadsheet will provide the underlying glue.

=back

=head1 SEE ALSO

L<DBD::CSV>

=head1 REPOSITORY

The public repository of this module is
L<https://github.com/Corion/DBIx-Spreadsheet>.

=head1 SUPPORT

The public support forum of this module is
L<https://perlmonks.org/>.

=head1 BUG TRACKER

Please report bugs in this module via the RT CPAN bug queue at
L<https://rt.cpan.org/Public/Dist/Display.html?Name=DBIx-Spreadsheet>
or via mail to L<dbix-spreadsheet-Bugs@rt.cpan.org>.

=head1 AUTHOR

Max Maischein C<corion@cpan.org>

=head1 COPYRIGHT (c)

Copyright 2019 by Max Maischein C<corion@cpan.org>.

=head1 LICENSE

This module is released under the same terms as Perl itself.

=cut
