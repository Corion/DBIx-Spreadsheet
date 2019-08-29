#!perl
use strict;
use DBI;
use Getopt::Long;
use Spreadsheet::Read;
use DBIx::RunSQL;
use Text::Unidecode;

# use File::Notify::Simple;
GetOptions();

my $file = $ARGV[0];
my $query = $ARGV[1];
warn "Reading '$file'";
my $data = read_file( $file );

sub read_file {
    # Make this smarter, later
    Spreadsheet::Read->new( $_[0] );
};

our $table_000;

sub gen_colnames {
    my %seen;
    my $i = 1;
    return map { qq{"$_"}}
           map { s!\s+!_!g; s!\W!!g; $_ }
           map { $i++; my $name = $_ eq '' ? sprintf "col_%d", $i : $seen{ $_ } ? "${_}_1" : $_; $seen{$name}++; $name }
           map { !defined($_) ? "" : $_ }
           @_;
}

sub import_data {
    my( $book ) = @_;
    my $dbh = DBI->connect('dbi:SQLite:dbname=:memory:',undef,undef,{AutoCommit => 1, RaiseError => 1,PrintError => 0});
    $dbh->sqlite_create_module(perl => "DBD::SQLite::VirtualTable::PerlData");

    my @tables;
    my $i = 0;
    for my $table_name ($book->sheets) {
        my $sheet = $book->sheet( $table_name );
        my $tablevar = sprintf 'table_%03d', $i++;
        my $data = [$sheet->rows()];
        my $colnames = shift @{$data};

        (my $sql_name = $table_name) =~ s!\s!_!g;

        # Fix up duplicate columns, empty column names

        $colnames = join ",", gen_colnames( @$colnames );
        {;
            #no strict 'refs';
            # Later, find the first non-empty row, instead of blindly taking the first row
            #${main::}{$tablevar} = \$data;
        };
        local $table_000 = $data;
        $tablevar = 'table_000';
        my $sql = qq(CREATE VIRTUAL TABLE temp."$sql_name" USING perl($colnames, arrayrefs="main::$tablevar"););
        $dbh->do($sql);
        push @tables, $sql_name;
    };

    return $dbh, \@tables;
}

my ($dbh, $tables) = import_data( $data );

DBIx::RunSQL->run(
    dbh => $dbh,
    sql => \$query,
);
