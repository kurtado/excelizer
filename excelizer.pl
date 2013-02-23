#!/usr/bin/perl
use strict;
use Data::Dumper;
use DBI;
use Text::CSV;
use MIME::Lite;
use POSIX qw(strftime);

# CONFIGURATIONS
# insert your favorite query here:
my $sql = "select * from emp";

# set email from / to
my $from = 'me@my.com';
my $to = 'you@your.com';

# set some db vars
my $dbtype = 'mysql';
my $dbname = 'dbname';
my $username = 'user';
my $pw = 'pass';

# EXECUTION
# get db handle and cursor
my $dbh = DBI->connect("dbi:$dbtype:dbname=$dbname", $username, $pw)
    or die "Couldn't connect to database: " . DBI->errstr;
my $sth = $dbh->prepare($sql)
    or die "Couldn't prepare query: " . $dbh->errstr;

# init csv 
my $csv = Text::CSV->new ( { binary => 1 } )
    or die "Cannot use CSV: ".Text::CSV->error_diag ();

# execute query
$sth->execute or die "Couldn't execute query: " . $sth->errstr;

# Print out the first row, which is a list of the columns in the select
my $mailstr = '';
$mailstr = $csv->string($csv->combine(@{ $sth->{NAME_lc} })) . "\n";

# main data loop
while (my $row = $sth->fetchrow_arrayref) {
    $mailstr .= $csv->string($csv->combine(@{$row})) . "\n";
}
#print "$mailstr\n";

# cleanup
$sth->finish;
$dbh->disconnect;

# get time stamp
my $now_string = strftime "%a %b %e %H:%M:%S %Y", localtime;

# mail it
my $msg = MIME::Lite->new(
    From    => $from,
    To      =>  $to,
#    Cc      => 'some@other.com, some@more.com',
    Subject => 'Excelizer Data',
    Type    => 'multipart/mixed'
    );

# attach email text and excelized data
$msg->attach(
    Type     => 'TEXT',
    Data     => "Excelizer data sent at $now_string"
    );
$msg->attach(
    Type     => 'application/vnd.ms-excel',
    Data => $mailstr,
    Filename => 'excelizer.xls',
    Disposition => 'attachment'
    );

### use Net:SMTP to do the sending
$msg->send('smtp','localhost');

