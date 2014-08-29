#!/usr/bin/perl

# BIG O FORTUNA
# Big O (https://en.wikipedia.org/wiki/Big_O_notation) & O Fortuna (https://en.wikipedia.org/wiki/O_Fortuna)
# Cryptocurrency analytics
# Developed by Josh Smith © 2013

use strict;
#use warnings;
use Spreadsheet::ParseExcel;
use Spreadsheet::WriteExcel;

use LWP::Simple;
use WWW::Mechanize;
use POSIX qw( strftime );

sub run{
	while(1){
		&main;
		sleep(1800);
	}
	print "bigofortuna complete\n";
}

sub main{
	my $filename = 'data.xls';
	my @worksheetlist = ('bitcoin', 'litecoin', 'peercoin', 'namecoin', 'novacoin');

	my %latesttransactions = ();

	foreach my $currency (@worksheetlist){
		$latesttransactions{$currency} = '';
	}

	my $parser   = new Spreadsheet::ParseExcel;
	my $readbook;
	my %oldworkbook;

	# See what exists
	if(-e "$filename"){
	    print "$filename exists\n";
	    # Pull what exists
	    $readbook = $parser->Parse("$filename");

	    # Parse what exists
		for my $worksheet ( $readbook->worksheets() ) {
			my @thisWorksheet;
			my @attributes;

			my $worksheetname = $worksheet->get_name();

		    my ( $row_min, $row_max ) = $worksheet->row_range();
		    my ( $col_min, $col_max ) = $worksheet->col_range();

		    for my $row ( $row_min .. $row_max ) {
		    	if($row == 0){
					for my $col ( $col_min .. $col_max ) {
			            my $cell = $worksheet->get_cell( $row, $col );

			            if(defined $cell){
			            	my $val = $cell->value();
			            	push(@attributes, $val);
			      		}
			        }
		    	} else {
		    		my %thisTransaction = ();

			        for my $col ( $col_min .. $col_max ) {
			            my $cell = $worksheet->get_cell( $row, $col );

			            if(defined $cell){
			            	my $val = $cell->value();

			            	unless("$attributes[$col]" eq "Transaction ID"){
			            		$thisTransaction{ $attributes[$col] } = $val;
			            	}

			            	if("$attributes[$col]" eq "DateTime"){
			            		if($val gt $latesttransactions{$worksheetname}){
			            			$latesttransactions{$worksheetname} = $val;
			            		}
			            	}

			            } else {
			            	$thisTransaction{ $attributes[$col] } = "";
			            }
			        }
			        push(@thisWorksheet, \%thisTransaction);
		    	}
		    }
		    $oldworkbook{ $worksheetname } = \@thisWorksheet;
		}

	} else {
	    print "$filename does not exist\n";
	}


	# GET new data

	# Bitcoin
	my $url = 'http://www.bitstamp.net/api/transactions/';
	my $mech = WWW::Mechanize->new();
	$mech->get($url);

	# Parse new data
	my $decodedtransactions = $mech->response()->decoded_content();

	my @dirtyArray = split('{', $decodedtransactions);

	my @newBitcoinTransactions;

	my %attributeTitles = (
	    tid => 'Transaction ID',
	    date => 'DateTime',
	    price => 'Price',
	    amount => 'Amount'
	);

	foreach my $transaction (@dirtyArray) {
		my %thisTransaction = ();
		unless ("$transaction" eq "[") {
			my @transactionAttributes = split(',', $transaction);
			foreach my $attribute (@transactionAttributes) {
				$attribute =~ s/[\[' }"}\]]//g;
				my @tempAttributeArray = split(':', $attribute);
				my $attributeName = $tempAttributeArray[0];
				my $attributeValue = $tempAttributeArray[1];
				if("$attributeName" eq "date"){
					$attributeValue = strftime("%Y-%m-%d %H:%M:%S", localtime($attributeValue));
				}
				$thisTransaction{ "$attributeTitles{$attributeName}" } = "$attributeValue";
			}
		}
		unless("$thisTransaction{'DateTime'}" le "$latesttransactions{'bitcoin'}"){
			push(@newBitcoinTransactions, \%thisTransaction);
		}
	}


	# Combine old transactions and new transactions
	my @alltransactions = @newBitcoinTransactions;

	if(%oldworkbook){
		my @oldtransactions = @{$oldworkbook{ 'bitcoin' }};
		@alltransactions = (@oldtransactions, @newBitcoinTransactions);	
	}

	@alltransactions = sort { $b->{'DateTime'} cmp $a->{'DateTime'} } @alltransactions;


	# Write all transactions to file
	print "overwriting $filename...\n";
	my $newbook = Spreadsheet::WriteExcel->new("$filename");

	foreach my $newworksheet (@worksheetlist){
		my $newsheet = $newbook->add_worksheet($newworksheet);

		$newsheet->write(0,0, 'DateTime');
		$newsheet->write(0,1, 'Price');
		$newsheet->write(0,2, 'Amount');

		if("$newworksheet" eq "bitcoin"){
			my $counter = 1;
			foreach my $writetransaction (@alltransactions){
				my %transhash = %{ $writetransaction };

				my $transhash_date = $transhash{'DateTime'};
				my $transhash_price = $transhash{'Price'};
				my $transhash_amount = $transhash{'Amount'};

				$newsheet->write($counter,0, $transhash_date);
				$newsheet->write($counter,1, $transhash_price);
				$newsheet->write($counter,2, $transhash_amount);

				$counter++;

			}
		}
	}
	my $time = localtime;
	print "$filename overwritten at $time\n";
}

&run;

