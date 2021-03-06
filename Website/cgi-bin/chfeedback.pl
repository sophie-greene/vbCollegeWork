#!/usr/bin/perl

##!/usr/bin/perl
#    chfeedback.pl Feedback Form Perl Script Ver 2.2.2
#    Generated by thesitewizard.com's Feedback Form Wizard.
#    Copyright 2000-2006 by Christopher Heng. All rights reserved.
#
#    thesitewizard and thefreecountry are trademarks of Christopher Heng.
#    $Id: perlscript.txt,v 1.5 2006/02/27 13:36:26 developer Exp $
#
#    Get the latest version, free, from:
#        http://www.thesitewizard.com/wizards/feedbackform.shtml
#
#	You can read the Frequently Asked Questions (FAQ) at:
#		http://www.thesitewizard.com/wizards/faq.shtml
#	
#	I can be contacted at:
#		http://www.thesitewizard.com/feedback.php
#	Note that I do not normally respond to questions that have
#	already been answered in the FAQ, so *please* read the FAQ.
#
#
#    LICENCE TERMS
#    
#    1. You may use this script on your website, with or
#    without modifications, free of charge.
#    
#    2. You may NOT distribute or republish this script, whether
#    modified or not. The script can only be distributed by the author,
#    Christopher Heng.
#    
#    3. THE SCRIPT AND ITS DOCUMENTATION ARE PROVIDED
#    "AS IS", WITHOUT WARRANTY OF ANY KIND, NOT EVEN THE
#    IMPLIED WARRANTY OF MECHANTABILITY OR FITNESS FOR A
#    PARTICULAR PURPOSE. YOU AGREE TO BEAR ALL RISKS AND
#    LIABILITIES ARISING FROM THE USE OF THE SCRIPT,
#    ITS DOCUMENTATION AND THE INFORMATION PROVIDED BY THE
#    SCRIPTS AND THE DOCUMENTATION.
#
#    If you cannot agree to any of the above conditions, you
#    may not use the script. 
#    
#    Although it is NOT required, I would be most grateful
#    if you could also link to us at:
#
#       http://www.thesitewizard.com/
#


# ---------- USER CONFIGURATION SECTION ---------------
# Before this script will do anything useful, the following
# variables must be set.
#
# MANDATORY VARIABLES
# $mailprog - the location of your mail program and the parameters
#           to pass to it.
#			eg $mailprog = "/usr/lib/sendmail" ;
# $mailto - email address where the feedback will be sent
#           eg, $mailto = 'yourname@example.com' ;
# $subject - the subject line in the email sent by the feedback form
#           eg $subject = "Feedback Form" ;
#
# $formurl - the URL of your feedback form
#           eg $formurl = "http://www.example.com/feedback.html" ;
# $thankyouurl - the URL of your thank you page
#           eg $thankyouurl = "http://www.example.com/thanks.html" ;
# $errorurl - the URL of your error page
#           eg $errorurl = "http://www.example.com/error.html" ;

$mailprog = "http://epub//usr/lib/sendmail" ;
$mailto = 'somoud_ss@hotmail.com' ;
$subject = "Feedback Form" ;
$formurl = "http://epub/062643/feedback.html" ;
$errorurl = "http://epub/062643/feedback.html" ;
$thankyouurl = "http://epub/062643/feedbackthankyou.html" ;

# ---------- END OF USER CONFIGURATION SECTION ------------

# ---------- functions -----------
sub redirect_url {
	my ( $url )	= shift ;
	print "Location: $url\n\n" ;
}

sub parse_form_data {
	my ($request_method, $input_string, $content_length) ;
	my (@vars, $indiv_var, $name, $value);
	my (%form) ;

	$request_method	= $ENV{'REQUEST_METHOD'} ;
	$content_length	= $ENV{'CONTENT_LENGTH'} ;

	# load the entire string into $input_string
	if ($request_method	=~ /post/i) {
		$input_string	= "" ;
		read( STDIN, $input_string, $content_length ) ;
	}
	else {
		$input_string	= $ENV{ 'QUERY_STRING' };
	}
	unless (defined $input_string) {
		$input_string	= "" ;
	}

	# put all the variable pairs (name=value) into an array
	@vars	= split( /&/, $input_string );

	# process each individual name, value pair, putting them
	# into a hash for easy access
	foreach $indiv_var ( @vars ) {

		# separate the pair
		($name, $value)	= split( /=/, $indiv_var, 2 );

		# translate encoding
		$name	=~ s/%([\da-fA-F]{2})/pack("C", hex($1))/eg;
		$name	=~ tr/+/ /;
		unless (defined $value) {
			# just in case there was no equals as in ISINDEX
			$value	= "" ;
		}
		$value	=~ s/%([\da-fA-F]{2})/pack("C", hex($1))/eg;
		$value	=~ tr/+/ /;

		# put the pair in the hash for easy access
		$form{$name}	= $value ;
	}

	return wantarray ? %form : undef ;
}

sub send_email {
	my ($mailprog, $email, $name, $mailto, $subject, $message) = @_ ;

	if (open MAIL, "|$mailprog -t") {
		print MAIL "To: $mailto\n" ;
		print MAIL "From: \"$name\" <$email>\n" ;
		print MAIL "Reply-To: \"$name\" <$email>\n" ;
		print MAIL "X-Mailer: chfeedback.pl 2.2.2\n" ;
		print MAIL "Subject: $subject\n\n" ;
		print MAIL $message ;
		close MAIL ;
	}
	# ignore fails, just don't do anything

	return ;
}

# ----------- main program ------

my %form = parse_form_data();

my $email ;
if (exists $form{"email"}) {
	$email	= $form{"email"} ;
}
else {
	redirect_url( $formurl );
	exit ;
}
if ($email eq "") {
	redirect_url( $errorurl );
	exit ;
}

my $name = $form{"name"} ;
if ($name eq "") {
	redirect_url( $errorurl );
	exit ;
}

my $comments = $form{'comments'} ;
if ($comments eq "") {
	redirect_url( $errorurl );
	exit ;
}

$name	=~ s/[\r\n].*//s;
$email	=~ s/[\r\n].*//s;
my $http_referer = $ENV{ 'HTTP_REFERER' };
my $message	=
	"This message was sent from:\n" .
	"$http_referer\n" .
	"------------------------------------------------------------\n" .
	"Name of sender: " . $name . "\n" .
	"Email of sender: " . $email . "\n" .
	"------------------------- COMMENTS -------------------------\n\n" .
	$comments .
	"\n\n------------------------------------------------------------\n" ;

send_email ( $mailprog, $email, $name, $mailto, $subject, $message );
redirect_url( $thankyouurl );
