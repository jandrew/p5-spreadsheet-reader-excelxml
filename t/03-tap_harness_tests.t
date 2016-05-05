#!/usr/bin/env perl
my	$dir 	= './';
my	$tests	= 'Spreadsheet/Reader/';
my	$up		= '../';
for my $next ( <*> ){
	if( ($next eq 't') and -d $next ){
		$dir	= './t/';
		$up		= '';
		last;
	}
}

use	TAP::Formatter::Console;
my $formatter = TAP::Formatter::Console->new({
					jobs => 1,
					#~ verbosity => 1,
				});
my	$args ={
		lib =>[
			$up . 'lib',
			$up,
			#~ $up . '../Log-Shiras/lib',
		],
		test_args =>{
			load_test					=>[],
			pod_test					=>[],
			error_test					=>[],
			cell_to_column_row_test		=>[],
			stacked_flag_test			=>[],
			cell_test					=>[],
			xml_workbook_rels_test		=>[],
			no_pivot_test				=>[ $dir . 'test_files/' ],
			temp_dir_test				=>[ $dir . 'test_files/' ],
			open_by_worksheet_test		=>[ $dir . 'test_files/' ],
			has_chart_test				=>[ $dir . 'test_files/' ],
			workbook_test				=>[ $dir . 'test_files/' ],
			types_test					=>[ $dir . 'test_files/' ],
			empty_sharedstrings_test	=>[ $dir . 'test_files/' ],
			shared_strings_test			=>[ $dir . 'test_files/' ],
			percent_file_test			=>[ $dir . 'test_files/' ],
			hidden_formatting_test		=>[ $dir . 'test_files/' ],
			losing_is_hidden_test		=>[ $dir . 'test_files/' ],
			merged_areas_test			=>[ $dir . 'test_files/' ],
			read_xlsm_feature			=>[ $dir . 'test_files/' ],
			text_in_worksheet_test		=>[ $dir . 'test_files/' ],
			open_xml_files				=>[ $dir . 'test_files/' ],
			quote_in_styles_test		=>[ $dir . 'test_files/' ],
			open_MySQL_files			=>[ $dir . 'test_files/' ],
			dxfId_handling_test			=>[ $dir . 'test_files/' ],
			generic_reader_test			=>[ $dir . 'test_files/' ],
			named_shared_strings_test	=>[ $dir . 'test_files/' ],
			named_styles_test			=>[ $dir . 'test_files/' ],
			position_styles_test		=>[ $dir . 'test_files/' ],
			named_worksheet_test		=>[ $dir . 'test_files/' ],
			worksheet_to_row_test		=>[ $dir . 'test_files/' ],
			worksheet_test				=>[ $dir . 'test_files/' ],
			lower_workbook_test			=>[ $dir . 'test_files/' ],
			zip_reader_test				=>[ $dir . 'test_files/' ],
			workbook_file_test			=>[ $dir . 'test_files/' ],
			xml_workbook_meta_test		=>[ $dir . 'test_files/' ],
			xml_workbook_props_test		=>[ $dir . 'test_files/' ],
			workbook_integration_test	=>[ $dir . 'test_files/' ],
			zip_workbook_meta_test		=>[ $dir . 'test_files/xl/' ],
			styles_sheet_test			=>[ $dir . 'test_files/xl/' ],
			position_shared_strings_test	=>[ $dir . 'test_files/xl/' ],
			shared_strings_interface_test	=>[ $dir . 'test_files/xl/' ],
			file_worksheet_test				=>[ $dir . 'test_files/xl/worksheets/' ],
			zip_workbook_rels_test			=>[ $dir . 'test_files/xl/_rels/' ],
			zip_workbook_props_test			=>[ $dir . 'test_files/docProps/' ],
		},
		formatter => $formatter,
	};
my	@tests =(
		[  $dir . '01-load.t', 'load_test' ],
		[  $dir . '02-pod.t', 'pod_test' ],
		[  $dir . $tests . 'ExcelXML/01-types.t', 'types_test' ],
		[  $dir . $tests . 'ExcelXML/02-error.t', 'error_test' ],
		[  $dir . $tests . 'ExcelXML/04-xml_reader.t', 'generic_reader_test' ],
		[  $dir . $tests . 'ExcelXML/05-cell_to_column_row.t', 'cell_to_column_row_test' ],
		[  $dir . $tests . 'ExcelXML/XMLReader/01-position_shared_strings.t', 'position_shared_strings_test' ],
		[  $dir . $tests . 'ExcelXML/XMLReader/02-named_shared_strings.t', 'named_shared_strings_test' ],
		[  $dir . $tests . 'ExcelXML/06-sharedstrings_interface.t', 'shared_strings_interface_test' ],
		[  $dir . $tests . 'ExcelXML/14-empty_sharedstrings.t', 'empty_sharedstrings_test' ],
		[  $dir . $tests . 'ExcelXML/XMLReader/03-position_styles.t', 'position_styles_test' ],
		[  $dir . $tests . 'ExcelXML/XMLReader/04-named_styles.t', 'named_styles_test' ],
		[  $dir . $tests . 'ExcelXML/07-styles_interface.t', 'styles_sheet_test' ],
		[  $dir . $tests . 'ExcelXML/16-quote_in_style_line.t', 'quote_in_styles_test' ],
		[  $dir . $tests . 'ExcelXML/17-dxfId_handling.t', 'dxfId_handling_test' ],
		[  $dir . $tests . 'ExcelXML/08-cell.t', 'cell_test' ],
		[  $dir . $tests . 'ExcelXML/XMLReader/05-file_worksheet.t', 'file_worksheet_test' ],
		[  $dir . $tests . 'ExcelXML/XMLReader/06-named_worksheet.t', 'named_worksheet_test' ],
		[  $dir . $tests . 'ExcelXML/09-worksheet_to_row.t', 'worksheet_to_row_test' ],
		[  $dir . $tests . 'ExcelXML/10-worksheet.t', 'worksheet_test' ],
		[  $dir . $tests . 'ExcelXML/15-text_in_worksheet.t', 'text_in_worksheet_test' ],
		[  $dir . $tests . 'ExcelXML/11-zip_reader.t', 'zip_reader_test' ],
		[  $dir . $tests . 'ExcelXML/12-workbook_file_interface.t', 'workbook_file_test' ],
		[  $dir . $tests . 'ExcelXML/ZipReader/01-zip_workbook_meta.t', 'zip_workbook_meta_test' ],
		[  $dir . $tests . 'ExcelXML/XMLReader/07-xml_workbook_meta.t', 'xml_workbook_meta_test' ],
		[  $dir . $tests . 'ExcelXML/ZipReader/02-zip_workbook_rels.t', 'zip_workbook_rels_test' ],
		[  $dir . $tests . 'ExcelXML/XMLReader/08-xml_workbook_rels.t', 'xml_workbook_rels_test' ],
		[  $dir . $tests . 'ExcelXML/ZipReader/03-zip_workbook_props.t', 'zip_workbook_props_test' ],
		[  $dir . $tests . 'ExcelXML/XMLReader/09-xml_workbook_props.t', 'xml_workbook_props_test' ],
		[  $dir . $tests . 'ExcelXML/13-workbook_integration.t', 'workbook_integration_test' ],
		[  $dir . $tests . '01-excelxml.t', 'workbook_test' ],
		[  $dir . $tests . '02-open_by_worksheet.t', 'open_by_worksheet_test' ],
		[  $dir . $tests . '03-temp_dir.t', 'temp_dir_test' ],
		[  $dir . $tests . '04-no_pivot.t', 'no_pivot_test' ],
		[  $dir . $tests . '05-chart.t', 'has_chart_test' ],
		[  $dir . $tests . '06-stacked_flags.t', 'stacked_flag_test' ],
		[  $dir . $tests . '07-losing_is_hidden.t', 'losing_is_hidden_test' ],
		[  $dir . $tests . '08-open_spreadsheet_ml_files.t', 'open_xml_files' ],
		[  $dir . $tests . '09-open_MySQL_data.t', 'open_MySQL_files' ],
		[  $dir . $tests . '10-shared_strings.t', 'shared_strings_test' ],
		[  $dir . $tests . '11-percent_file.t', 'percent_file_test' ],
		[  $dir . $tests . '12-merge_function_alignment.t', 'merged_areas_test' ],
		[  $dir . $tests . '13-hidden_formatting.t', 'hidden_formatting_test' ],
		[  $dir . $tests . '14-read_xlsm_feature.t', 'read_xlsm_feature' ],
	);
use	TAP::Harness;
use	TAP::Parser::Aggregator;
my	$harness	= TAP::Harness->new( $args );
my	$aggregator	= TAP::Parser::Aggregator->new;
	$aggregator->start();
	$harness->aggregate_tests( $aggregator, @tests );
	$aggregator->stop();
use Test::More;
explain $formatter->summary($aggregator);
pass( "Test Harness Testing complete" );
done_testing();