package Spreadsheet::Reader::ExcelXML::XMLReader;
use version; our $VERSION = version->declare('v0.1_1');
###LogSD	warn "You uncovered internal logging statements for Spreadsheet::Reader::ExcelXML::XMLReader-$VERSION";

use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use Try::Tiny;
use Text::ParseWords 3.27;
use Types::Standard qw(
		Int				Bool				Enum
		Num				Str					ArrayRef
    );
use Carp 'confess';
#~ use Data::Dumper;
use lib	'../../../../lib',;
###LogSD	with 'Log::Shiras::LogSpace';
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use Spreadsheet::Reader::ExcelXML::Types qw( IOFileType );
use	Spreadsheet::Reader::ExcelXML::Error;

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has file =>(
		isa			=> IOFileType,
		reader		=> 'get_file',
		writer		=> 'set_file',
		predicate	=> 'has_file',
		clearer		=> 'clear_file',
		coerce		=> 1,
		trigger		=> \&_start_xml_reader,
		handles 	=> [qw( close getline )],
	);

has workbook_inst =>(
		isa	=> 'Spreadsheet::Reader::ExcelXML::Workbook', 
		handles =>[ qw(
			set_error						get_empty_return_type			_get_workbook_file_type
			_get_sheet_info					_get_rel_info					get_sheet_names
			get_defined_conversion			set_defined_excel_formats		has_shared_strings_interface
			get_shared_string				get_values_only					is_empty_the_end
			_starts_at_the_edge				get_group_return_type			change_output_encoding
			counting_from_zero				get_error_inst					boundary_flag_setting
			has_styles_interface			get_format						start_the_ss_file_over
			parse_excel_format_string		has_error						get_default_format
			_get_workbook_file_type
		)],
		writer => 'set_workbook_inst',
		predicate => '_has_workbook_inst',
	);

has	xml_version =>(
		isa			=> 	Num,
		reader		=> 'version',
		writer		=> '_set_xml_version',
		clearer		=> '_clear_xml_version',
		predicate	=> '_has_xml_version',
	);

has	xml_encoding =>(
		isa			=> 	Str,
		reader		=> 'encoding',
		predicate	=> 'has_encoding',
		writer		=> '_set_xml_encoding',
		clearer		=> '_clear_xml_encoding',
	);

has	xml_header =>(
		isa			=> 	Str,
		reader		=> 'get_header',
		writer		=> '_set_xml_header',
	);

has position_index =>(
		isa			=> Int,
		reader		=> 'where_am_i',
		writer		=> 'i_am_here',
		clearer		=> 'clear_location',
		predicate	=> 'has_position',
	);

has	file_type =>(
		isa			=> 	Str,
		reader		=> 'get_file_type',
		default		=> 'xml',
	);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9


#~ sub start_the_file_over{
	#~ my( $self ) = @_;
	#~ ###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	#~ ###LogSD			$self->get_all_space . '::XMLReader::start_the_file_over', );
	#~ $self->clear_location;
	#~ ###LogSD	$phone->talk( level => 'debug', message =>[ "location is cleared" ] );
	#~ if( $self->_has_xml_parser ){
		#~ $self->_close_the_sheet;
		#~ ###LogSD	$phone->talk( level => 'debug', message =>[ "sheet is closed" ] );
		#~ $self->_clear_xml_parser;
		#~ ###LogSD	$phone->talk( level => 'debug', message =>[ "parser is cleared" ] );
	#~ }
	#~ if( $self->has_file ){
		#~ ###LogSD		$phone->talk( level => 'debug', message =>[ "Resetting the XML file" ] );
		#~ my $fh = $self->get_file;
		#~ ###LogSD		$phone->talk( level => 'debug', message =>[ "got the file handle", $fh ] );
		#~ $fh->seek( 0, 0 );
		#~ ###LogSD		$phone->talk( level => 'debug', message =>[ "seek to 0 done" ] );
		#~ my $xml_parser = XML::LibXML::Reader->new( IO => $fh );
		#~ ###LogSD		$phone->talk( level => 'debug', message =>[ "XML Parser built" ] );
		
		#~ # Check functionality
		#~ if( eval '$xml_parser->read' ){
			#~ ###LogSD	$phone->talk( level => 'debug', message =>[
			#~ ###LogSD		'Initial read successful - re-loading the reader' ], );
		#~ }else{
			#~ ###LogSD	$phone->talk( level => 'debug', message =>[
			#~ ###LogSD		'Initial read failed - clearing the reader' ], );
			#~ $self->set_error( $@ );
			#~ $xml_parser = undef;
		#~ }
		
		#~ # Load to class
		#~ if( $xml_parser ){
			#~ $self->_set_xml_parser( $xml_parser );
			#~ ###LogSD	$phone->talk( level => 'debug', message =>[ "XML parser stored", $xml_parser ] );
		#~ }else{
			#~ ###LogSD	$phone->talk( level => 'debug', message =>[ "Failed to build an xml parser" ] );
			#~ return undef;
		#~ }
		#~ return 1;
	#~ }else{
		#~ ###LogSD		$phone->talk( level => 'info', message =>[ "No file to reset" ] );
		#~ return undef;
	#~ }
#~ }

sub get_text_node{
	my ( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::XMLReader::get_text_node', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"checking the text value of the node", $self->location_status, $self->_node_value ] );
	
	# Check for a text node type (and return immediatly if so)
	if( $self->_has_value ){
		my $node_text = $self->_node_value;
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"This is a text node - returning value |$node_text|",] );
		return ( 1, $node_text, );
	}
	# Return undef for no value
	return ( undef,);
}

sub get_attribute_hash_ref{
	my ( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::XMLReader::get_attribute_hash_ref', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Extract all attributes as a hash ref", ] );
	
	my $attribute_ref = {};
	my $result = $self->_move_to_first_att;
	###LogSD	$phone->talk( level => 'trace', message =>[
	###LogSD		"Result of the first attribute move: $result",] );
	ATTRIBUTELIST: while( $result > 0 ){
		my $att_name = $self->_node_name;
		my $att_value = $self->_node_value;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Reading attribute: $att_name", "..and value: $att_value" ] );
		if( $att_name eq 'val' ){
			$attribute_ref = $att_value;
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Assuming we are at the bottom of the attribute list with a found attribute val: $att_value"] );
			last ATTRIBUTELIST;
		}else{
			$attribute_ref->{$att_name} = "$att_value";
		}
		$result = $self->_move_to_next_att;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Result of the move: $result", ] );
	}
	$result = ( ref $attribute_ref ) ? (keys %$attribute_ref) : 1;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Returning attribute ref:", $attribute_ref ] );
	return ( $result, $attribute_ref );
}

sub advance_element_position{
	my ( $self, $element, $position ) = @_;
	if( $position and $position < 1 ){
		confess "You can only advance element position in a positive direction, |$position| is not correct.";
	}
	$position ||= 1;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::XMLReader::advance_element_position', );
	###LogSD	$phone->talk( level => 'info', message => [
	###LogSD		"Advancing to element -" . ($element//'') . "- -$position- times", ] );
	
	#~ # Check for end of file and opt out
	my( $node_depth, $node_name, $node_type ) = $self->location_status;
	#~ if( $node_name eq 'EOF' ){
		#~ ###LogSD	$phone->talk( level => 'debug', message => [
		#~ ###LogSD		"Already at the EOF - returning failure", ] );
		#~ return undef ;
	#~ }
	
	my $result;
	my $x = 0;
	for my $y ( 1 .. $position ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Advancing position iteration: $y", ] );
		if( defined $element ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Searching for element: $element", "with location status:", $self->location_status ] );
			eval '$result = $self->_next_element( $element )';#eval ''
			if( $@ ){
				###LogSD	$phone->talk( level => 'fatal', message => [
				###LogSD		"_next_element failed with: $@", ] );
				return;
			}
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"search result: " . ($result//'none'), ] );
		}else{
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Generic node indexing", ] );
			( $result, $node_depth, $node_name, $node_type ) = $self->_next_node;
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Received the result: $result", "..at depth: $node_depth",
			###LogSD		"..and node named: $node_name", "..of node type: $node_type" ] );
			
			# Climb out of end tags
			while( $result and $node_type == 15 ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Advancing from end node", ] );
				( $result, $node_depth, $node_name, $node_type ) = $self->_next_node;
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Received the result: $result", "..at depth: $node_depth",
				###LogSD		"..and node named: $node_name", "..of node type: $node_type" ] );
			}
		}
		last if !$result;
		$x++;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Successfully indexed -$x- times", ] );
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Successfully indexed forward a total -$position- times", ] );
	
	if( defined $node_type and $node_type == 0 ){
		###LogSD	$phone->talk( level => 'info', message =>[ "Reached the end of the file!" ] );
	}elsif( !$result ){
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"Unable to location position -$position- for element: " . ($element//'') ] );
	}else{
		###LogSD	$phone->talk( level => 'info', message => [
		###LogSD		"Actually advanced -$x- positions with result: $result",
		###LogSD		"..indicated by:", $self->location_status ] );
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"returning result: " . ($result//'none'), ] );
	return $result;
}

sub location_status{
	my ( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::XMLReader::location_status', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Getting the status for the current position" ] );#, , caller(0), caller(1), caller(2), caller(3)
	eval '$self->get_node_all'; # Fixes a phantom xml declaration added at the end of the reader without being in the file
	if( $@ ){
		###LogSD	$phone->talk( level => 'warn', message => [
		###LogSD		"Couldn't get the node because: " . $@ ] );#,
		$self->set_error( $@ );
	}
	my ( $node_depth, $node_name, $node_type ) = ( $self->_node_depth, $self->_node_name, $self->_node_type );
	$node_name	= 
		( $node_type == 0 ) ? 'EOF' :
		( $node_name eq '#text') ? 'raw_text' :
		$node_name;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Currently at libxml2 level: $node_depth",
	###LogSD		"Current node name: $node_name",
	###LogSD		"..for type: $node_type" ] );
	return ( $node_depth, $node_name, $node_type );
}
	

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has _node_stack =>(
	isa		=> ArrayRef,
	traits	=> ['Array'],
	reader	=> '_get_node_stack',
	writer	=> '_set_node_stack',
	default	=> sub{ [] },
	handles	=>{
		_add_node		=> 'unshift',# Handle nodes in a fifo rather than lifo order
		_remove_node	=> 'pop',
		_has_nodes		=> 'count',
		_current_node	=>[ get => -1 ],
		_process_node	=>[ get => 0 ],
	}
);

#~ has _xml_reader =>(
	#~ isa			=> 'XML::LibXML::Reader',
	#~ reader		=> '_get_xml_parser',
	#~ writer		=> '_set_xml_parser',
	#~ predicate	=> '_has_xml_parser',
	#~ clearer		=> '_clear_xml_parser',
	#~ handles	=>{
		#~ _close_the_sheet	=> 'close',
		#~ _node_depth			=> 'depth',
		#~ _node_type			=> 'nodeType',
		#~ _node_name			=> 'name',
		#~ _encoding			=> 'encoding',
		#~ _version			=> 'xmlVersion',
		#~ _next_element		=> 'nextElement',
		#~ _node_value			=> 'value',
		#~ _has_value			=> 'hasValue',
		#~ _move_to_first_att	=> 'moveToFirstAttribute',
		#~ _move_to_next_att	=> 'moveToNextAttribute',
		#~ _read_next_node		=> 'read',
		#~ skip_siblings		=> 'skipSiblings',
		#~ next_sibling		=> 'nextSibling',
		#~ get_node_all		=> 'readOuterXml',
	#~ },
#~ );

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _start_xml_reader{
	my( $self, $file_handle ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::XMLReader::_start_xml_reader', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"turning a file handle into an xml reader", ] );
	
	$self->clear_location;
	###LogSD	$phone->talk( level => 'debug', message =>[ "location is cleared" ] );
	
	# Set defaults for the file as needed
	###LogSD	$phone->talk( level => 'debug', message =>[ "start at the beginning" ] );
	$file_handle->seek( 0, 0 );
	###LogSD		$phone->talk( level => 'trace', message => [
	###LogSD			"made it past seek (to the beginning of the file)", ] );
	$self->_next_node;
	# Set the file unique bits
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		"Check if this type of file has unique settings" ], );
	if( $self->can( '_load_unique_bits' ) ){
		###LogSD	$phone->talk( level => 'debug', message =>[ "Loading unique bits" ], );
		$self->_load_unique_bits;
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Finished loading unique bits" 			], );
	}
	###LogSD	$phone->talk( level => 'debug', message => [ "file built" ], );
	$self->start_the_file_over if $self->has_file;
	
		#~ }else{
			#~ ###LogSD	$phone->talk( level => 'debug', message =>[
			#~ ###LogSD		'Initial read failed - clearing the reader' ], );
			#~ $self->set_error( $@ );
			#~ $xml_reader = undef;
		#~ }
	#~ }
	#~ if( !$xml_reader ){
		#~ ###LogSD	$phone->talk( level => 'debug', message =>[ 'Not dead yet - setting error: ' . $@ ], );
		#~ $self->set_error( $@ ) if !$self->has_error;
		#~ $self->clear_file;
		#~ $self->clear_location;
		#~ ###LogSD	$phone->talk( level => 'debug', message =>[ "location is cleared" ] );
		#~ $self->_clear_xml_parser;
	#~ }
	###LogSD	$phone->talk( level => 'debug', message => [ "finished all xml reader build steps" ], );
}

sub _next_node{
	my( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::XMLReader::_next_node', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"reading the next node", ] );
	
	# Handle a previously read node
	if( $self->_has_nodes ){
		###LogSD	$phone->talk( level => 'debug', message =>[ "Removed the old node", ] );
		$self->_remove_node;
	}
	
	# Get the next node
	my $node;
	if( !$self->_has_nodes ){
		$self->_read_file;
		###LogSD	$phone->talk( level => 'debug', message =>[ "Need to add nodes to the stack" ] );
	}
	$node = $self->_current_node;
	###LogSD	$phone->talk( level => 'debug', message =>[ "Stored node is:", $node] );
	
	# Check if clarification is needed farther in the file
	
	# Return what is needed
	if( wantarray ){
		my( $node_depth, $node_name, $node_type ) = 
			( $self->_get_current_level, $self->_get_current_name, $self->_get_current_type );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Returning a succefull result: 1", "..at depth: $node_depth",
		###LogSD		"..to node named: $node_name", "..and node type: $node_type" ] );
		return( 1, $node_depth, $node_name, $node_type );
	}else{
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Returning a succesfull result: 1", ] );
		return 1;
	}
	
	
}

sub _read_file{
	my( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::XMLReader::_read_file', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"reading the 'line' from the file", ] );
		
	# get the next file line
	my @sections;
	while( !@sections ){
		my $line = $self->getline; # Check for end of file here!!!!!!!!!!!
		###LogSD	$phone->talk( level => 'debug', message =>[ "The next file line is: " . $line ] );
		$line = substr( $line, 0, -1 );
		@sections = split />/, $line;
		###LogSD	$phone->talk( level => 'debug', message =>[ "Line sections are:", @sections ] );
		if( scalar( @sections ) > 2 ){
			confess 'The next xml line broke into more than 2 sections |' . join( '~|~', @sections ) . '| from line: ' . $line;
		}
	}
	
	# Pull name, type, and attributes as well as calculating depth
	my @types = qw( regular raw_text );
	my @closures = qw( open closed );
	my $is_xml_header = 0;
	my $x = 0;
	for my $node ( @sections ){
		###LogSD	$phone->talk( level => 'debug', message =>[ "Processing section: ", $node] );
		my $initial_string = $node;
		
		# Handle xml escapes
		$node =~ s/&lt;/</g;
		$node =~ s/&gt;/>/g;
		$node =~ s/&quot;/"/g;
		$node =~ s/&amp;/&/g;
		$node =~ s/&apos;/'/g;
		$node =~ s/&#60;/</g;
		$node =~ s/&#62;/>/g;
		$node =~ s/&#34;/"/g;
		$node =~ s/&#38;/&/g;
		$node =~ s/&#39;/'/g;
		
		# Handle header nodes with quotes
		if( substr( $node, 0, 1 ) eq '?' ){
			$is_xml_header = 1;
			$node = substr( $node, 1, -1 );
			###LogSD	$phone->talk( level => 'debug', message =>[ "Removed question marks from node: " . $node ] );
		}
		
		my $node_ref;
		# Handle start and terminus nodes
		if( substr( $node, 0, 1 ) eq '/' ){
			###LogSD	$phone->talk( level => 'fatal', message =>[
			###LogSD		"Reached an end node - should probably index the level and re-call this method?" ] );
		}elsif( $is_xml_header ){
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"xml headers are closed node levels by definition" ] );
			$node_ref->{closed} = 'closed';
		}elsif( substr( $node, -1, 1 ) eq '/' ){ ###### Potentially add a node name stack here to validate the level?
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Self closing node found" ] );
			$node = substr( $node, 0, -1 );
			$node_ref->{closed} = 'closed';
		}else{
			$node_ref->{closed} = $closures[$x];
		}
		###LogSD	$phone->talk( level => 'debug', message =>[ "New node ref:", $node_ref ] );
		
		# Set node level
		if( $self->_has_nodes ){
			my $node = $self->_current_node;
			$node_ref->{level} = $node->{level} + ($node->{closed} eq 'closed' ? 0 : 1);
		}else{
			$node_ref->{level} = 0;
		}
		###LogSD	$phone->talk( level => 'debug', message =>[ "updated node ref:", $node_ref ] );
		
		# Split out the node name and the attributes
		my @attributes = parse_line( '\s+', 0, $node );
		###LogSD	$phone->talk( level => 'debug', message =>[ "Words after split: ", @attributes] );
		$node_ref->{name} = shift @attributes;
		$node_ref->{type} = $types[$x++];
		###LogSD	$phone->talk( level => 'debug', message =>[ "updated node ref:", $node_ref ] );
		
		# Store the xml header string
		$self->_set_xml_header( $initial_string ) if $is_xml_header and $node_ref->{name} eq 'xml';
		
		# Parse the attributes
		for my $att ( @attributes ){
			my( $att_name, $att_val ) = split /\s*=\s*/, $att;
			$att_val = substr( $att_val, 1, -3 ) if substr( $att_val, 0, 1 ) eq '"';
			$node_ref->{attributes}->{$att_name} = $att_val;
			###LogSD	$phone->talk( level => 'debug', message =>[ "updated node ref:", $node_ref ] );
		}
		
		# Store the node
		$self->_add_node( $node_ref );
		last if $node_ref->{closed} eq 'closed';
	}
	
	# Handle (and clear) header nodes then move on
	if( $is_xml_header ){
		$self->_load_header;
		$self->_read_file;
	}
	
	# If you made it here the process worked
	return 1;
}

sub _load_header{
	my( $self, $reader ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::XMLReader::_load_header', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"loading file level settings since the header was found", ] );
	
	my $header_node = $self->_remove_node;
	for my $attribute ( qw( version encoding progid ) ){
		if( exists $header_node->{attributes}->{$attribute} ){
			my $setter = '_set_xml_' . $attribute;
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Performing the action -$setter- on the data: $header_node->{attributes}->{$attribute}", ] );
			$self->$setter( $header_node->{attributes}->{$attribute} );
		}
	}
}
			
	

#~ sub _reader_init{
	#~ my( $self, $reader ) = @_;
	#~ ###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	#~ ###LogSD			$self->get_all_space . '::XMLReader::_reader_init', );
	#~ ###LogSD		$phone->talk( level => 'debug', message => [
	#~ ###LogSD			"loading any file specific settings", ] );
		
	#~ # Set basic xml values
	#~ my	$xml_string = '<?xml version="';
	#~ $self->_next_node;
	#~ if( $self->_version ){
		#~ $self->_set_xml_version( $self->_version );
		#~ $xml_string .= $self->_version . '"';
	#~ }else{
		#~ confess "Could not find the version of this xml document!";
	#~ }
	#~ if( $self->_encoding ){
		#~ $self->_set_xml_encoding( $self->_encoding );
		#~ $xml_string .= ' encoding="' . $self->_encoding . '"'
	#~ }else{
		#~ $self->_clear_xml_encoding;
	#~ }
	#~ $xml_string .= '?>';
	#~ $self->_set_xml_header( $xml_string );
	#~ ###LogSD	$phone->talk( level => 'debug', message => [ "Finished the base first pass - file initialization", $xml_string ], );
#~ }

#~ sub _next_node{
	#~ my( $self, ) = @_;
	#~ ###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	#~ ###LogSD			$self->get_all_space . '::XMLReader::_next_node', );
	#~ ###LogSD		$phone->talk( level => 'debug', message => [
	#~ ###LogSD			"Reading the next node in the xml document"] );
	
	#~ # Attempt to read the node
	#~ my $result = 0;
	#~ if( !eval '$self->_read_next_node' ){
		#~ ###LogSD	$phone->talk( level => 'debug', message => [
		#~ ###LogSD		"found a node read error: $@",] );
		#~ $self->set_error( $@ );
		#~ my $fh = $self->get_file;
		#~ $fh->seek(0,0);
		#~ ###LogSD	$phone->talk( level => 'debug', message =>[ "got the file handle", $fh ] );
		#~ map{ chomp;
		#~ ###LogSD	$phone->talk( level => 'debug', message =>[ $_ ] );
		#~ } <$fh>;
	#~ }else{
		#~ $result = 1;
	#~ }
	
	#~ # Check for an unexpected end of the document
	#~ my ( $node_depth, $node_name, $node_type ) = $self->location_status;
	#~ if( $node_name eq '#document' and $node_depth == 0 ){
		#~ ###LogSD	$phone->talk( level => 'debug', message => [
		#~ ###LogSD		"Reached the unexpected end of the document", ] );
		#~ $result = 0;
	#~ }
	
	#~ if( wantarray ){
		#~ ###LogSD	$phone->talk( level => 'debug', message => [
		#~ ###LogSD		"Returning the result: $result", "..at depth: $node_depth",
		#~ ###LogSD		"..to node named: $node_name", "..and node type: $node_type" ] );
		#~ return( $result, $node_depth, $node_name, $node_type );
	#~ }else{
		#~ ###LogSD	$phone->talk( level => 'debug', message => [
		#~ ###LogSD		"Returning the result: $result", ] );
		#~ return $result;
	#~ }
#~ }

sub _close_file_and_reader{
	my ( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::XMLReader::DEMOLISH', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"clearing the XMLReader reader for log space:",
	###LogSD			$self->get_all_space . '::XMLReader::DEMOLISH', ] );
	
	# Close the file
	if( $self->has_file ){
		###LogSD	$phone->talk( level => 'debug', message =>[ "Closing the file handle", ] );
		$self->clear_file
	}
	#~ print "XMLReader file check complete\n";
	
}

sub DEMOLISH{
	my ( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::_hidden::DEMOLISH', );
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD			"XMLReader DEMOLISH called" ] );
	
	$self->_close_file_and_reader;
	
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::Reader::ExcelXML::XMLReader - A LibXML::Reader xlsx base class

=head1 SYNOPSIS

	package MyPackage;
	use MooseX::StrictConstructor;
	use MooseX::HasDefaults::RO;
	extends	'Spreadsheet::Reader::ExcelXML::XMLReader';
    
=head1 DESCRIPTION

This documentation is written to explain ways to use this module when writing your own excel 
parser.  To use the general package for excel parsing out of the box please review the 
documentation for L<Workbooks|Spreadsheet::Reader::ExcelXML>,
L<Worksheets|Spreadsheet::Reader::ExcelXML::Worksheet>, and 
L<Cells|Spreadsheet::Reader::ExcelXML::Cell>

This module provides a generic way to open an xml file or xml file handle and then extract 
information using the L<XML::LibXML::Reader> parser.  The additional methods and attributes 
are intended to provide some coalated parsing commands that are specifically useful in turning 
xml to perl data structures.

=head2 Attributes

Data passed to new when creating an instance.  For modification of these attributes see the 
listed 'attribute methods'. For general information on attributes see 
L<Moose::Manual::Attributes>.  For ways to manage the instance when opened see the 
L<Methods|/Methods>.
	
=head3 file

=over

B<Definition:> This attribute holds the file handle for the file being read.  If the full 
file name and path is passed to the attribute it is coerced to an IO::File file handle.

B<Default:> no default - this must be provided to read a file

B<Required:> yes

B<Range:> any unencrypted xml file name and path or IO::File file handle

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<set_file>

=over

B<Definition:> change the file value in the attribute (this will reboot 
the file instance and lock the file)

=back

B<get_file>

=over

B<Definition:> Returns the file handle of the file even if a file name 
was passed

=back

B<has_file>

=over

B<Definition:> this is used to see if the file loaded correctly.

=back

B<clear_file>

=over

B<Definition:> this clears (and unlocks) the file handle

=back

=back

L<Delegated Methods>

=over

B<close>

=over

closes the file handle

=back

=back

=back

=head3 workbook_inst

=over

B<Definition:> This attribute holds a reference to the top level workbook (parser).  
The purpose is to use some of the attributes stored there.

B<Default:> no default

B<Required:> not strictly but methods will fail without the L<delegated
|/Delegated Methods (required)> elements listed below.  For information 
about the delegated methods see the links to the documentation.  For 
hidden methods in indicator of where to look in the code is provded

B<Range:> isa => 'Spreadsheet::Reader::ExcelXML'

B<weak_ref> => 1

B<attribute methods> Methods provided to adjust this attribute

=over

B<set_workbook_inst>

=over

set the attribute with a workbook instance

=back

=back

B<Delegated Methods (required)> Methods delegated to this module by the attribute
		
=over

error, set_error, clear_error => L<Spreadsheet::Reader::ExcelXML/error_inst>

get_empty_return_type => L<Spreadsheet::Reader::ExcelXML/empty_return_type>

get_sheet_name => L<Spreadsheet::Reader::ExcelXML/get_sheet_name>

sheet_count => L<Spreadsheet::Reader::ExcelXML/sheet_count>

_get_sheet_info, _get_rel_info, _get_id_info, _set_sheet_info, 
_get_workbook_file_type, _get_sheet_lookup => all hidden methods that accesss 
the hidden attribute _workbook_meta_data in the parser which delegates from 
L<Spreadsheet::Reader::ExcelXML::WorkbookMetaInterface>

=back

=back

=head3 xml_version

=over

B<Definition:> This stores the xml version stored in the xml header.  It is read 
when the file handle is first set in this sheet.

B<Default:> no default - this is auto read from the header

B<Required:> no

B<Range:> xml versions

B<attribute methods> Methods provided to adjust this attribute

=over

B<version>

=over

get the stored xml version

=back

=back

=over

B<_clear_xml_version>

=over

clear the attribute value

=back

=back

B<_set_xml_version>

=over

set the attribute value

=back

=back

=head3 xml_encoding

=over

B<Definition:> This stores the data encoding of the xml file from the xml header.  
It is read when the file handle is first set in this sheet.

B<Default:> no default - this is auto read from the header

B<Required:> no

B<Range:> valid xml file encoding

B<attribute methods> Methods provided to adjust this attribute

=over

B<encoding>

=over

get the attribute value

=back

=back

=over

B<has_encoding>

=over

predicate for the attribute value

=back

=back

=over

B<_clear_xml_encoding>

=over

clear the attribute value

=back

=back

B<_set_xml_encoding>

=over

set the attribute value

=back

=back

=head3 xml_header

=over

B<Definition:> This stores the xml header from the xml file.  It is read when 
the file handle is first set in this sheet.  I contains both the verion and 
the encoding where available

B<Default:> no default - this is auto read from the header

B<Required:> no

B<Range:> valid xml file header

B<attribute methods> Methods provided to adjust this attribute

=over

B<get_header>

=over

get the attribute value

=back

=back

=over

B<_set_xml_header>

=over

set the attribute value

=back

=back

=back

=head3 position_index

=over

B<Definition:> This attribute is available to facilitate other consuming roles and 
classes.  Of the attribute methods only the 'clear_location' method is used in this 
class during the 'start_the_file_over' method.  It can be used for tracking same level 
positions with the same node name.

B<Default:> no default - this is mostly managed by the child class or add on role

B<Required:> no

B<Range:> Integer

B<attribute methods> Methods provided to adjust this attribute

=over

B<where_am_i>

=over

get the attribute value

=back

=back

=over

B<i_am_here>

=over

set the attribute value

=back

=back

=over

B<clear_location>

=over

clear the attribute value

=back

=back

=over

B<has_position>

=over

set the attribute value

=back

=back

=back

=head3 file_type

=over

B<Definition:> This is a static attribute that shows the file type

B<Default:> xml

B<attribute methods> Methods provided to adjust this attribute

=over

B<get_file_type>

=over

get the attribute value

=back

=back

=back

=head2 Methods

These are the methods provided by this class only.  They do not incude any methods added 
by roles to this class elsewhere

=head3 start_the_file_over

=over

B<Definition:> This will disconnect the L<XML::LibXML::Reader> from the file handle,  
rewind the file handle, and then reconnect the L<XML::LibXML::Reader> to the file handle.

B<Accepts:> nothing

B<Returns:> nothing

=back

=head3 get_text_node

=over

B<Definition:> This will collect the text node at the current node position.  It will return 
two items ( $success_or_failure, $text_node_value )

B<Accepts:> nothing

B<Returns:> ( $success_or_failure(1|undef), ($text_node_value|undef) )

=back

=head3 get_attribute_hash_ref

=over

B<Definition:> Some nodes have attribute settings.  This method returns a hashref with any 
attribute settings attached as key => value pairs or an empty hash for no attributes

B<Accepts:> nothing

B<Returns:> { attribute_1 => attribute_1_value ... etc. }

=back

=head3 advance_element_position( [$node_name], [$number_of_times_to_index] )

=over

B<Definition:> This method will attempt to advance to $node_name (optional) or the next node 
if no $node_name is passed.  If there is an expectation of multiple nodes of the same name at 
the same level you can also pass $number_of_times_to_index (optional).  This will move through 
the xml file at the $node_name level the number of times indicated starting with wherever the 
xml file is already located.  Meaning $number_of_times_to_index is a relative index not an 
absolute index.

B<Accepts:> nothing

B<Returns:> success or failure for the method call

=back

=head3 location_status

=over

B<Definition:> This method gives three usefull location values with one call

B<Accepts:> nothing

B<Returns:> ( $node_depth (from the top of the file), $node_name, $node_type (xml numerical value for type) );

=back

=head2 Delegated Methods

These are the methods delegated to this class from L<XML::LibXML::Reader>.  For more 
general parsing of subsections of the xml file also see L<Spreadsheet::Reader::ExcelXML>.

=head3 skip_siblings => L<XML::LibXML::Reader/skipSiblings ()>

=head3 copy_current_node => L<XML::LibXML::Reader/readOuterXml ()>

=head1 SUPPORT

=over

L<github Spreadsheet::Reader::ExcelXML/issues
|https://github.com/jandrew/Spreadsheet-XLSX-Reader-LibXML/issues>

=back

=head1 TODO

=over

B<1.> Nothing currently

=back

=head1 AUTHOR

=over

=item Jed Lund

=item jandrew@cpan.org

=back

=head1 COPYRIGHT

This program is free software; you can redistribute
it and/or modify it under the same terms as Perl itself.

The full text of the license can be found in the
LICENSE file included with this module.

This software is copyrighted (c) 2014, 2015 by Jed Lund

=head1 DEPENDENCIES

=over

L<XML::LibXML::Reader>

=back

=head1 SEE ALSO

=over

L<Spreadsheet::ParseExcel> - Excel 2003 and earlier

L<Spreadsheet::XLSX> - 2007+

L<Spreadsheet::ParseXLSX> - 2007+

L<Log::Shiras|https://github.com/jandrew/Log-Shiras>

=over

All lines in this package that use Log::Shiras are commented out

=back

=back

=cut

#########1#########2 main pod documentation end   5#########6#########7#########8#########9