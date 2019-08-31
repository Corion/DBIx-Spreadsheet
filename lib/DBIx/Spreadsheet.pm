#!perl
package DBIx::Spreadsheet;
use strict;
use DBI;
use Getopt::Long;
use Spreadsheet::Read;
use Text::Unidecode;

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

=cut

has 'file' => (
    is => 'ro',
);

has 'spreadsheet' => (
    is => 'lazy',
    default => \&_read_file,
);

has 'dbh' => (
    is => 'lazy',
    default => sub( $self ) { $self->_import_data; $self->{dbh} },
);

has 'tables' => (
    is => 'lazy',
    default => sub( $self ) { $self->_import_data; $self->{tables} },
);

has 'spreadsheet_options' => (
    is => 'lazy',
    default => sub { {
        dtfmt => 'yyyy-mm-dd',
    } },
);

sub _read_file( $self ) {
    Spreadsheet::Read->new( $self->file, %{ $self->spreadsheet_options } );
};

our $table_000;

sub gen_colnames( $self, @colnames ) {
    my %seen;
    my $i = 1;
    return map { qq{"$_"}}
           map { s!\s+!_!g; s!\W!!g; $_ }
           map { $i++; my $name = $_ eq '' ? sprintf "col_%d", $i : $seen{ $_ } ? "${_}_1" : $_; $seen{$name}++; $name }
           map { !defined($_) ? "" : $_ }
           @colnames;
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
        my $data = [map { [$sheet->cellrow($_)] } 1..$sheet->maxrow ];
        #my $data = [$sheet->rows($_)];
        my $colnames = shift @{$data};

        (my $sql_name = $table_name) =~ s!\s!_!g;

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
