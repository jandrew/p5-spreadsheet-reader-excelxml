#########1 Test File for Spreadsheet::Reader::ExcelXML  6#########7#########8#########9

my ( $lib, $test_file );
BEGIN{
	$ENV{PERL_TYPE_TINY_XS} = 0;
	my	$start_deeper = 1;
	$lib		= 'lib';
	$test_file	= 't/test_files/';
	for my $next ( <*> ){
		if( ($next eq 't') and -d $next ){
			$start_deeper = 0;
			last;
		}
	}
	if( $start_deeper ){
		$lib		= '../../../' . $lib;
		$test_file	= '../../test_files/'
	}
	use Carp 'longmess';
	$SIG{__WARN__} = sub{ print longmess $_[0]; $_[0]; };
}
$| = 1;

use	Test::Most tests => 29;
use	Test::Moose;
use Data::Dumper;
use	lib	'../../../../Log-Shiras/lib',
		'../../../../MooseX-ShortCut-BuildInstance/lib',
		$lib,
	;
#~ use Log::Shiras::Switchboard v0.21 qw( :debug );#
###LogSD	my	$operator = Log::Shiras::Switchboard->get_operator(
###LogSD			name_space_bounds =>{
###LogSD				UNBLOCK =>{
###LogSD					log_file => 'warn',
###LogSD				},
###LogSD				build_instance =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'warn',
###LogSD					},
###LogSD				},
###LogSD				build_class =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'warn',
###LogSD					},
###LogSD				},
###LogSD				Test =>{
#~ ###LogSD					SharedStringsInstance =>{
#~ ###LogSD						UNBLOCK =>{
#~ ###LogSD							log_file => 'trace',
#~ ###LogSD						},
#~ ###LogSD					},
#~ ###LogSD					StylesInstance =>{
#~ ###LogSD						UNBLOCK =>{
#~ ###LogSD							log_file => 'warn',
#~ ###LogSD						},
#~ ###LogSD					},
#~ ###LogSD					WorkbookFileInterface =>{
#~ ###LogSD						UNBLOCK =>{
#~ ###LogSD							log_file => 'warn',
#~ ###LogSD						},
#~ ###LogSD					},
#~ ###LogSD					WorkbookMetaInterface =>{
#~ ###LogSD						UNBLOCK =>{
#~ ###LogSD							log_file => 'trace',
#~ ###LogSD						},
#~ ###LogSD						XMLReader=>{
#~ ###LogSD							_hidden =>{
#~ ###LogSD								UNBLOCK =>{
#~ ###LogSD									log_file => 'warn',
#~ ###LogSD								},
#~ ###LogSD							},
#~ ###LogSD							squash_node =>{
#~ ###LogSD								UNBLOCK =>{
#~ ###LogSD									log_file => 'warn',
#~ ###LogSD								},
#~ ###LogSD							},
#~ ###LogSD							current_named_node =>{
#~ ###LogSD								UNBLOCK =>{
#~ ###LogSD									log_file => 'warn',
#~ ###LogSD								},
#~ ###LogSD							},
#~ ###LogSD						},
#~ ###LogSD					},
#~ ###LogSD					WorkbookRelsInterface =>{
#~ ###LogSD						UNBLOCK =>{
#~ ###LogSD							log_file => 'warn',
#~ ###LogSD						},
#~ ###LogSD					},
#~ ###LogSD					WorkbookPropsInterface =>{
#~ ###LogSD						UNBLOCK =>{
#~ ###LogSD							log_file => 'warn',
#~ ###LogSD						},
#~ ###LogSD					},
#~ ###LogSD					SharedStringsInterface =>{
#~ ###LogSD						UNBLOCK =>{
#~ ###LogSD							log_file => 'warn',
#~ ###LogSD						},
#~ ###LogSD					},
#~ ###LogSD					StylesInterface =>{
#~ ###LogSD						UNBLOCK =>{
#~ ###LogSD							log_file => 'warn',
#~ ###LogSD						},
#~ ###LogSD					},
#~ ###LogSD					ExcelFormatInterface =>{
#~ ###LogSD						UNBLOCK =>{
#~ ###LogSD							log_file => 'warn',
#~ ###LogSD						},
#~ ###LogSD					},
#~ ###LogSD					Top =>{
#~ ###LogSD						_hidden =>{
#~ ###LogSD							BUILDARGS =>{
#~ ###LogSD								UNBLOCK =>{
#~ ###LogSD									log_file => 'warn',
#~ ###LogSD								},
#~ ###LogSD							},
#~ ###LogSD						},
#~ ###LogSD						_around =>{
#~ ###LogSD							parse =>{
#~ ###LogSD								UNBLOCK =>{
#~ ###LogSD									log_file => 'debug',
#~ ###LogSD								},
#~ ###LogSD							},
#~ ###LogSD						},
#~ ###LogSD					},
#~ ###LogSD					Workbook =>{
#~ ###LogSD						_hidden =>{
#~ ###LogSD							BUILDARGS =>{
#~ ###LogSD								UNBLOCK =>{
#~ ###LogSD									log_file => 'warn',
#~ ###LogSD								},
#~ ###LogSD							},
#~ ###LogSD							BUILD =>{
#~ ###LogSD								UNBLOCK =>{
#~ ###LogSD									log_file => 'warn',
#~ ###LogSD								},
#~ ###LogSD							},
#~ ###LogSD							set_formatter_inst =>{
#~ ###LogSD								UNBLOCK =>{
#~ ###LogSD									log_file => 'warn',
#~ ###LogSD								},
#~ ###LogSD							},
#~ ###LogSD							_build_workbook =>{
#~ ###LogSD								UNBLOCK =>{
#~ ###LogSD									log_file => 'warn',
#~ ###LogSD								},
#~ ###LogSD							},
#~ ###LogSD							_build_file_interface =>{
#~ ###LogSD								UNBLOCK =>{
#~ ###LogSD									log_file => 'warn',
#~ ###LogSD								},
#~ ###LogSD							},
#~ ###LogSD						},
#~ ###LogSD					},
###LogSD				},
###LogSD			},
###LogSD			reports =>{
###LogSD				log_file =>[ Print::Log->new ],
###LogSD			},
###LogSD		);
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
###LogSD	use MooseX::ShortCut::BuildInstance;
use Spreadsheet::Reader::ExcelXML;# ':debug'
#~ use Spreadsheet::Reader::ExcelXML::Error;
$test_file = ( @ARGV ) ? $ARGV[0] : $test_file;
$test_file .= 'values.xlsx';
	#~ print "Test file is: $test_file\n";
my  ( 
		$parser, @worksheets, $value,
	);
my	$answer_ref = [
		'Blad1',
		[0,5],
		[0,1],
		['Spreadsheet::Reader::ExcelXML::Cell', 'A1'],
		['Spreadsheet::Reader::ExcelXML::Cell', ' ',],
		['', undef,],
		['Spreadsheet::Reader::ExcelXML::Cell', 0,],
		['Spreadsheet::Reader::ExcelXML::Cell', 1,],
		['Spreadsheet::Reader::ExcelXML::Cell', '',],
		['Spreadsheet::Reader::ExcelXML::Cell', 'label'],
		['Spreadsheet::Reader::ExcelXML::Cell', 'space',],
		['Spreadsheet::Reader::ExcelXML::Cell', 'empty',],
		['Spreadsheet::Reader::ExcelXML::Cell', 'nul',],
		['Spreadsheet::Reader::ExcelXML::Cell', 'one',],
		['Spreadsheet::Reader::ExcelXML::Cell', 'quote',],					

	];
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'info', message => [ "harder questions ..." ] );
lives_ok{
			$parser = Spreadsheet::Reader::ExcelXML->new(
							###LogSD log_space => 'Test'
						);
			$parser->set_warnings( 0 );
			$parser->parse($test_file);
}										"Prep a test parser instance";
###LogSD		$phone->talk( level => 'info', message => [ "$parser:", $parser ] );
is			$parser->error(), undef,	"Write any error messages from the file load";
ok			@worksheets = $parser->worksheets(),
										"Loaded worksheet objects ok";
			my	$x = 0;
			for my $worksheet ( @worksheets ){
is			$worksheet->get_name, $answer_ref->[$x],
										'Check that the next opened worksheet name is: ' . $answer_ref->[$x++];
			my @column_range = $worksheet->col_range;
is_deeply	[@column_range], $answer_ref->[$x++],
										"Check for the correct column range";
			my @row_range = $worksheet->row_range;
is_deeply	[@row_range], $answer_ref->[$x++],
										"Check for the correct row range";
			for my $row ( $row_range[0] .. $row_range[1] ){
			for my $col ( $column_range[0] .. $column_range[1] ){
###LogSD	my $row_check = 0; my $column_check = 0;
###LogSD	if( $row == $row_check and $col == $column_check ){
###LogSD		$operator->add_name_space_bounds( {
#~ ###LogSD			Test =>{
###LogSD				UNBLOCK =>{
###LogSD					log_file => 'trace',
###LogSD				},
#~ ###LogSD			},
###LogSD		} );
###LogSD	}
#~ ###LogSD	elsif( $row > $row_check or ($row == $row_check and $col > $column_check ) ){
#~ ###LogSD		exit 1;
#~ ###LogSD	}
			my $cell;
is			ref( $cell = $worksheet->get_cell( $row, $col ) ), $answer_ref->[$x]->[0],
										"Attempt to get the cell for row -$row- and column -$col-";
			if( $answer_ref->[$x]->[0] ){
is			$cell->value, $answer_ref->[$x]->[1],
										"And check the returned value: " . $answer_ref->[$x]->[1];
			}
			$x++;
			#~ while( !$value or $value ne 'EOF' ){
#~ ok			$value = ($worksheet->get_next_value//'undef'),
										#~ "Get the next value position:$x";
#~ my			$return = ref( $value ) ? $value->value : $value;
#~ is			$return, $answer_ref->[$x],
										#~ "With value: " . $answer_ref->[$x];
#~ explain		$value_position;
			#~ $x++;
			}
			}
			}
#~ is_deeply	$parser->get_worksheet_names, $answer_ref->[$x++],
										#~ "Check that the overall worksheet list does not contains the chartsheet element";
#~ is			$parser->worksheet( 'Chart1' ), undef,
										#~ "Try to return the chartsheet: Chart1";
#~ like		$parser->error, $answer_ref->[$x],
										#~ "..and check for the correct error: " . $answer_ref->[$x++];
explain 								"...Test Done";
done_testing();

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
