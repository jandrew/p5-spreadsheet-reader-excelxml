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
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML.pm',
						"The Spreadsheet::Reader::ExcelXML file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/Cell.pm',
						"The Spreadsheet::Reader::ExcelXML::Cell file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/CellToColumnRow.pm',
						"The Spreadsheet::Reader::ExcelXML::CellToColumnRow file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/Chartsheet.pm',
						"The Spreadsheet::Reader::ExcelXML::Chartsheet file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/Error.pm',
						"The Spreadsheet::Reader::ExcelXML::Error file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/Row.pm',
						"The Spreadsheet::Reader::ExcelXML::Row file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/SharedStrings.pm',
						"The Spreadsheet::Reader::ExcelXML::SharedStrings file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/Styles.pm',
						"The Spreadsheet::Reader::ExcelXML::Styles file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/Types.pm',
						"The Spreadsheet::Reader::ExcelXML::Types file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/Workbook.pm',
						"The Spreadsheet::Reader::ExcelXML::Workbook file has good POD" );
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
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/XMLReader/FileWorksheet.pm',
						"The Spreadsheet::Reader::ExcelXML::XMLReader::FileWorksheet file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/XMLReader/NamedSharedStrings.pm',
						"The Spreadsheet::Reader::ExcelXML::XMLReader::NamedSharedStrings file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/XMLReader/NamedStyles.pm',
						"The Spreadsheet::Reader::ExcelXML::XMLReader::NamedStyles file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/XMLReader/NamedWorksheet.pm',
						"The Spreadsheet::Reader::ExcelXML::XMLReader::NamedWorksheet file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/XMLReader/PositionSharedStrings.pm',
						"The Spreadsheet::Reader::ExcelXML::XMLReader::PositionSharedStrings file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/XMLReader/PositionStyles.pm',
						"The Spreadsheet::Reader::ExcelXML::XMLReader::PositionStyles file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/XMLReader/WorkbookMeta.pm',
						"The Spreadsheet::Reader::ExcelXML::XMLReader::WorkbookMeta file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/XMLReader/WorkbookProps.pm',
						"The Spreadsheet::Reader::ExcelXML::XMLReader::WorkbookProps file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/XMLReader/WorkbookRels.pm',
						"The Spreadsheet::Reader::ExcelXML::XMLReader::WorkbookRels file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/XMLReader/WorkbookXML.pm',
						"The Spreadsheet::Reader::ExcelXML::XMLReader::WorkbookXML file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/ZipReader.pm',
						"The Spreadsheet::Reader::ExcelXML::ZipReader file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/ZipReader/WorkbookMeta.pm',
						"The Spreadsheet::Reader::ExcelXML::ZipReader::WorkbookMeta file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/ZipReader/WorkbookRels.pm',
						"The Spreadsheet::Reader::ExcelXML::ZipReader::WorkbookRels file has good POD" );
pod_file_ok( $up . 	'lib/Spreadsheet/Reader/ExcelXML/ZipReader/WorkbookProps.pm',
						"The Spreadsheet::Reader::ExcelXML::ZipReader::WorkbookProps file has good POD" );
done_testing();