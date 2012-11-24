

sub AUTOLOAD()
{
	print "hello from the AUTOLOAD subroutine!\n";
	print "called procedure:\n";
	print $AUTOLOAD;
	print "arguments:\n";
	foreach my $elem (@_){
		print "$elem\n";
	
	}
	
	
	


}


NOTEXIST(6,3);