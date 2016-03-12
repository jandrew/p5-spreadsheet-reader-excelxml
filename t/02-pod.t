#!/usr/bin/env perl
### Test that the pod files run
use Test::More;
use Test::Pod 1.48;
my	$up		= '../';
for my $next ( <*> ){
	if( ($next eq 't') and -d $next ){
		### <where> - found the t directory - must be using prove ...
		$up	= '';
		last;
	}
}
pod_file_ok( $up . 	'README.pod',
						"The README file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/Cell.pm',
						"The Spreadsheet::Reader::ExcelXML::Cell file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/CellToColumnRow.pm',
						"The Spreadsheet::Reader::ExcelXML::CellToColumnRow file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/Chartsheet.pm',
						"The Spreadsheet::Reader::ExcelXML::Chartsheet file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/Error.pm',
						"The Spreadsheet::Reader::ExcelXML::Error file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/FmtDefault.pm',
						"The Spreadsheet::Reader::ExcelXML::FmtDefault file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/FormatInterface.pm',
						"The Spreadsheet::Reader::ExcelXML::FormatInterface file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/ParseExcelFormatStrings.pm',
						"The Spreadsheet::Reader::ExcelXML::ParseExcelFormatStrings file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/Row.pm',
						"The Spreadsheet::Reader::ExcelXML::Row file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/SharedStrings.pm',
						"The Spreadsheet::Reader::ExcelXML::SharedStrings file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/Styles.pm',
						"The Spreadsheet::Reader::ExcelXML::Styles file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/Types.pm',
						"The Spreadsheet::Reader::ExcelXML::Types file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/WorkbookFileInterface.pm',
						"The Spreadsheet::Reader::ExcelXML::WorkbookFileInterface file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/WorkbookMetaInterface.pm',
						"The Spreadsheet::Reader::ExcelXML::WorkbookMetaInterface file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/WorkbookPropsInterface.pm',
						"The Spreadsheet::Reader::ExcelXML::WorkbookPropsInterface file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/WorkbookRelsInterface.pm',
						"The Spreadsheet::Reader::ExcelXML::WorkbookRelsInterface file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/Worksheet.pm',
						"The Spreadsheet::Reader::ExcelXML::Worksheet file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/WorksheetToRow.pm',
						"The Spreadsheet::Reader::ExcelXML::WorksheetToRow file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/XMLReader.pm',
						"The Spreadsheet::Reader::ExcelXML::XMLReader file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/XMLToPerlData.pm',
						"The Spreadsheet::Reader::ExcelXML::XMLToPerlData file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/XMLReader/ExtractFile.pm',
						"The Spreadsheet::Reader::ExcelXML::XMLReader::ExtractFile file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/XMLReader/NamedStyles.pm',
						"The Spreadsheet::Reader::ExcelXML::XMLReader::NamedStyles file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/XMLReader/PositionStyles.pm',
						"The Spreadsheet::Reader::ExcelXML::XMLReader::PositionStyles file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/ZipReader.pm',
						"The Spreadsheet::Reader::ExcelXML::ZipReader file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/ZipReader/ExtractFile.pm',
						"The Spreadsheet::Reader::ExcelXML::ZipReader::ExtractFile file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/Workbook.pm',
						"The Spreadsheet::Reader::ExcelXML::Workbook file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML.pm',
						"The Spreadsheet::Reader::ExcelXML file has good POD" );
done_testing();