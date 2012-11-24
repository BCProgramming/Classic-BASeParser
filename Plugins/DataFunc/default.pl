



#sub pmax {
#        my $max = shift(@_);
#        foreach $foo (@_) {
#            $max = $foo if $max < $foo;
#        }
#        return $max;
#    }

sub pjoin{
	#joins all items in the argument list.
	#kind of- gay.
	my $strreturn;
	foreach my $elem (@_)
	{
		$strreturn=$strreturn + $elem;
	
	
	}
	return $strreturn;


}

print (pjoin("HI","BYE"));