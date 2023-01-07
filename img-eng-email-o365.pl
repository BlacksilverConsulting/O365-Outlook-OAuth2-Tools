#!/usr/bin/perl

use English;
use strict;
use warnings;

use chilkat();
use File::Temp qw/ tempfile /;

use Log::LogLite;

# TMW Modules
use lib "/var/imaging/lib";
use TmwDB qw(:all);
use Tmw qw(:all);

=head1 NAME

img-eng-email-o365.pl -- E-Mail Import Message retriever for Outlook365+OAuth2

=head1 DESCRIPTION

Retrieve messages from POP3 accounts and drop them in /home/mailer/X
for further processing, where X is the cnf_Email_Import.cei_Sequence
of the source of the message.

Each message is deleted from the mailbox after retrieval to prevent
duplication.

Messages are processed by a different program, img-eng-email-proc.pl.

Unlike most other engines, this does not run as listener, and it is not
lauched by img-listloop.

=head1 INVOCATION

Invoked by a scheduled task (cron job) for user 'mailer'.

=head1 PARAMETERS

Mostly cnf_Email_Import, some img_Setting.

The first command line parameter must contain only digits, and match a
directory under /home/mailer where message files will be written.

That parameter is also used to retrieve OAuth2 configuration and tokens
from /home/mailer/.o365.

For example, if the parameter value is "23", then messages will be
written to files in /home/mailer/23/, the OAuth2 configuration will be
read from /home/mailer/.o365/23-config.json, and the token will be
kept in /home/mailer/.o365/23-token.json

=head1 OAUTH2

OAuth2 configuration is a beast. This program does not generate either of
the JSON files mentioned above. It will update the access_token upon
successful refresh, which is attempted every time the program runs.

=head1 CHILKAT

This program requires the chilkat utlity library. Deal.

=head1 VERSION

$Id$

=cut

my $_loglevel = getsysparm( "System", "Logging Level" ) || 6;
my $_log = Log::LogLite->new( '/var/log/imaging/mailer', $_loglevel )
    or die "Can't open Log::LogLite";
$_log->default_message("($PID): ");
my @_log_prefix = (
    'PANIC', 'FAILURE', 'CRITICAL', 'ERROR', 'WARNING', 'NOTICE',
    'INFO',  'DEBUG',   'DEBUG',    'DEBUG', 'DEBUG'
);

sub logx {
    my ( $message_level, $message_text ) = @_;
    $message_level ||= 6;
    $_log->write( $_log_prefix[$message_level] . ": " . $message_text,
        $message_level );
    print "DEBUG: $message_text\n"
        if $_loglevel > 6 and $message_level <= $_loglevel;
    return;
}

# Initialize chilkat library
my $glob = chilkat::CkGlobal->new();
my $success = $glob->UnlockBundle("");
if ($success != 1) {
    logx 1, $glob->lastErrorText();
    exit;
}

logx 10, "Starting $PROGRAM_NAME";

my $oauth2 = chilkat::CkOAuth2->new();

if (not length $ARGV[0] ) {
    print "Missing sequence parameter.\n";
    exit;
}

my $sequence = $ARGV[0];
if ( $sequence !~ /^[0-9]+$/ ) {
    # Match only digits, leave only integers
    print "Invalid sequence parameter.\n";
    exit;
}

my $o365_dir = '/home/mailer/.o365';
my $jsonConfig = chilkat::CkJsonObject->new();
my $config_file_name = "$o365_dir/$sequence-config.json";
$success = $jsonConfig->LoadFile($config_file_name);
if ($success != 1) {
    print "Failed to load config file [$config_file_name]\n";
    exit;
}
print "AAD config loaded from [$config_file_name]\n";

my $url_base = "https://login.microsoftonline.com/";
$oauth2->put_AuthorizationEndpoint( $url_base . $jsonConfig->stringOf("tenantId") . "/oauth2/v2.0/authorize" );
$oauth2->put_TokenEndpoint( $url_base . $jsonConfig->stringOf("tenantId") . "/oauth2/v2.0/token" );

$oauth2->put_ClientId($jsonConfig->stringOf("clientId"));
$oauth2->put_ClientSecret($jsonConfig->stringOf("clientSecret"));

my $jsonToken = chilkat::CkJsonObject->new();
my $token_file_name = "$o365_dir/$sequence-token.json";
$success = $jsonToken->LoadFile($token_file_name);
if ($success != 1) {
    logx 1, "Failed to load token file [$token_file_name]";
    exit;
}
logx 7, "OAuth2 token loaded from [$token_file_name]";

$oauth2->put_RefreshToken($jsonToken->stringOf("refresh_token"));

# Send the HTTP POST to refresh the access token..
$success = $oauth2->RefreshAccessToken();
if ($success != 1) {
    logx 1, $oauth2->lastErrorText();
    exit;
}
logx 7, "OAuth2 authorization refresh succeeded!";

# The response contains a new access token, but we must keep
# our existing refresh token for when we need to refresh again in the future.
$jsonToken->UpdateString("access_token",$oauth2->accessToken());

# Save the JSON to a file for future requests.
my $fac = chilkat::CkFileAccess->new();
$success = $fac->WriteEntireTextFile($token_file_name,$jsonToken->emit(),"utf-8",0);
if ($success != 1) {
    logx 1, $oauth2->lastErrorText();
    exit;
}
logx 7, "Saved updated access_token to [$token_file_name]";

my $mailman = chilkat::CkMailMan->new();

my ( $server, $user, $password, $auth_metnod, $options )
     = list grep { $sequence == $_->[5] } list st 'EmailImportRetrieval';

$mailman->put_MailHost($server);
$mailman->put_MailPort(995);
$mailman->put_PopSsl(1);


# Use your O365 email address here.
$mailman->put_PopUsername($user);

# When using OAuth2 authentication, leave the password empty.
$mailman->put_PopPassword("");

$mailman->put_OAuth2AccessToken($jsonToken->stringOf("access_token"));

# Make the TLS connection to the outlook.office365.com POP3 server.
$success = $mailman->Pop3Connect();
if ($success != 1) {
    logx 1, $mailman->lastErrorText();
    exit;
}
logx 9, "Connected to POP3 server";

# Authenticate using XOAUTH2
$success = $mailman->Pop3Authenticate();
if ($success != 1) {
    logx 1, $mailman->lastErrorText();
    exit;
}
logx 6, "Authenticated to POP3 server using OAuth";

# Find out how many emails are on the server..
my $numEmails = $mailman->CheckMail();
if ($numEmails < 0) {
    logx 1, $mailman->lastErrorText();
    exit;
}
logx 7, "About to download [$numEmails] messages";

# Copy the all email from the user's POP3 mailbox
# into a bundle object.  The email remains on the server.
# CopyMail is a reasonable choice for POP3 maildrops that don't have too many
# emails. For larger mail drops, one might download emails one at a time..
# bundle is a EmailBundle
my $bundle = $mailman->CopyMail();
if ($mailman->get_LastMethodSuccess() != 1) {
    logx 1, $mailman->lastErrorText();
    exit;
}
logx 9, "Download complete";

my $dirPath = "./tmp";

my $bundleIndex = 0;
my $numMessages = $bundle->get_MessageCount();
my $toBeDeleted = chilkat::CkStringArray->new();

while (($bundleIndex < $numMessages)) {
logx 7, "Parsing message $bundleIndex";

    my $email = $bundle->GetEmail($bundleIndex);

    my ($fh, $filename) = tempfile( DIR => "/home/mailer/$sequence", SUFFIX => '.msg', UNLINK => 0 );
    logx 7, "Writing [$filename]";
    print $fh $email->getMime();
    close $fh;
    my $mode = "0660";
    chmod oct($mode), $filename
        or logx 3, "Failed to chmod: $!";

    $toBeDeleted->Append( $email->uidl() );

    $bundleIndex = $bundleIndex + 1;
}

if ( $toBeDeleted->get_Count() ) {
    $success = $mailman->DeleteMultiple( $toBeDeleted );
    if ($success != 1) {
        logx 1, $mailman->lastErrorText();
        exit;
    }
    logx 7, "Deleted " . $toBeDeleted->get_Count() . " messages";
}

# End the POP3 session and close the connection to the server.
$success = $mailman->Pop3EndSession();
if ($success != 1) {
    logx 1, $mailman->lastErrorText();
    exit;
}
logx 10, "Ended POP3 session";

logx 6, "Finished.";
