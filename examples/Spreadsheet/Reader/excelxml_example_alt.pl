#!/usr/bin/env perl
use strict;
use warnings;
use lib 	'../../../lib',
			'../../../../Log-Shiras/lib';
$| = 1;
#~ use Log::Shiras::Switchboard qw( :debug );#
###LogSD	my	$operator = Log::Shiras::Switchboard->get_operator(
###LogSD						name_space_bounds => {
###LogSD							Base=>{
###LogSD								Worksheet =>{
###LogSD									UNBLOCK =>{
###LogSD										log_file => 'trace',
###LogSD									},
###LogSD									_build_xml_reader =>{
###LogSD										BLOCK =>{ log_file => 1 },
###LogSD									},
###LogSD									start_the_file_over =>{
###LogSD										BLOCK =>{ log_file => 1 },
###LogSD									},
###LogSD									_load_unique_bits =>{
###LogSD										BLOCK =>{ log_file => 1 },
###LogSD									},
###LogSD									_reader_init =>{
###LogSD										BLOCK =>{ log_file => 1 },
###LogSD									},
###LogSD									_parse_column_row =>{
###LogSD										BLOCK =>{ log_file => 1 },
###LogSD									},
###LogSD									parse_element =>{
###LogSD										BLOCK =>{ log_file => 1 },
###LogSD									},
###LogSD									XMLReader =>{
###LogSD										DEMOLISH =>{
###LogSD											BLOCK =>{ log_file => 1 },
###LogSD										},
###LogSD									},
###LogSD								},
###LogSD							},
###LogSD						},
###LogSD						reports =>{
###LogSD							log_file =>[ Print::Log->new ],
###LogSD						},
###LogSD					);
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use Spreadsheet::Reader::ExcelXML v0.1_1 qw( :just_the_data );# :alt_default

my $workbook  = Spreadsheet::Reader::ExcelXML->new(
					file => '../../../t/test_files/TestBook.xlsx',
					###LogSD	log_space => 'Base',
				);

if ( !$workbook->file_opened ) {
	die $workbook->error(), "\n";
}

for my $worksheet ( $workbook->worksheet ) {
	
	print "Reading worksheet named: " . $worksheet->get_name . "\n";
	print "..at position: " . $worksheet->position . "\n";
	
	while( 1 ){ 
		my $cell = $worksheet->get_next_value;
		print "Cell is: $cell\n";
		last if $cell eq 'EOF';
	}
	#~ last;# In order not to read all sheets
}

###LogSD	package Print::Log;
###LogSD	use Data::Dumper;
###LogSD	sub new{
###LogSD		bless {}, shift;
###LogSD	}
###LogSD	sub add_line{
###LogSD		shift;
###LogSD		my @input = ( ref $_[0]->{message} eq 'ARRAY' ) ? 
###LogSD						@{$_[0]->{message}} : $_[0]->{message};
###LogSD		my ( @print_list, @initial_list );
###LogSD		no warnings 'uninitialized';
###LogSD		for my $value ( @input ){
###LogSD			push @initial_list, (( ref $value ) ? Dumper( $value ) : $value );
###LogSD		}
###LogSD		for my $line ( @initial_list ){
###LogSD			$line =~ s/\n$//;
###LogSD			$line =~ s/\n/\n\t\t/g;
###LogSD			push @print_list, $line;
###LogSD		}
###LogSD		printf( "| level - %-6s | name_space - %-s\n| line  - %04d   | file_name  - %-s\n\t:(\t%s ):\n", 
###LogSD					$_[0]->{level}, $_[0]->{name_space},
###LogSD					$_[0]->{line}, $_[0]->{filename},
###LogSD					join( "\n\t\t", @print_list ) 	);
###LogSD		use warnings 'uninitialized';
###LogSD	}

###LogSD	1;