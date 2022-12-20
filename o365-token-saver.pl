#!/usr/bin/perl

use English; 
use strict;
use warnings;

use chilkat();

# TODO: POD

# Much of this is based on the chilkat example at
# https://www.example-code.com/perl/office365_oauth2_access_token.asp

my $glob = chilkat::CkGlobal->new();
my $success = $glob->UnlockBundle("");
if ($success != 1) {
    print $glob->lastErrorText() . "\r\n";
    exit;
}

my $oauth2 = chilkat::CkOAuth2->new();

my $listen_port = 3017;
$oauth2->put_ListenPort($listen_port);
$oauth2->put_AppCallbackUrl( "http://localhost:$listen_port/email" );

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

my $jsonConfig = chilkat::CkJsonObject->new();
my $config_file_name = "$sequence-config.json";
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

$oauth2->put_CodeChallenge(0);

my @scopes = (
    "openid", 
    "profile", 
    "offline_access", 
    "https://outlook.office365.com/SMTP.Send", 
    "https://outlook.office365.com/POP.AccessAsUser.All", 
    "https://outlook.office365.com/IMAP.AccessAsUser.All" 
    );

$oauth2->put_Scope( join ' ', @scopes );

# Begin the OAuth2 three-legged flow.  This returns a URL that should be loaded in a browser.
my $url = $oauth2->startAuth();
if ( $oauth2->get_LastMethodSuccess() != 1 ) {
    print $oauth2->lastErrorText() . "\r\n";
    exit;
}

# This is specific to Windows. Your OS may vary.
`start "" "$url"`;

# Now wait for the authorization.
# We'll wait for a max of 30 seconds.
my $numMsWaited = 0;
while ( ( $numMsWaited < 30000 ) and ( $oauth2->get_AuthFlowState() < 3 ) ) {
    $oauth2->SleepMs( 100 );
    $numMsWaited = $numMsWaited + 100;
}

# If there was no response from the browser within 30 seconds, then
# the AuthFlowState will be equal to 1 or 2.
# 1: Waiting for Redirect. The OAuth2 background thread is waiting to receive the redirect HTTP request from the browser.
# 2: Waiting for Final Response. The OAuth2 background thread is waiting for the final access token response.
# In that case, cancel the background task started in the call to StartAuth.
if ( $oauth2->get_AuthFlowState() < 3 ) {
    $oauth2->Cancel();
    print "No response from the browser!" . "\r\n";
    exit;
}

# Check the AuthFlowState to see if authorization was granted, denied, or if some error occurred
# The possible AuthFlowState values are:
# 3: Completed with Success. The OAuth2 flow has completed, the background thread exited, and the successful JSON response is available in AccessTokenResponse property.
# 4: Completed with Access Denied. The OAuth2 flow has completed, the background thread exited, and the error JSON is available in AccessTokenResponse property.
# 5: Failed Prior to Completion. The OAuth2 flow failed to complete, the background thread exited, and the error information is available in the FailureInfo property.
if ( $oauth2->get_AuthFlowState() == 5 ) {
    print "OAuth2 failed to complete." . "\r\n";
    print $oauth2->failureInfo() . "\r\n";
    exit;
}

if ( $oauth2->get_AuthFlowState() == 4 ) {
    print "OAuth2 authorization was denied." . "\r\n";
    print $oauth2->accessTokenResponse() . "\r\n";
    exit;
}

if ( $oauth2->get_AuthFlowState() != 3 ) {
    print "Unexpected AuthFlowState:" . $oauth2->get_AuthFlowState() . "\r\n";
    exit;
}

print "OAuth2 authorization granted!" . "\r\n";
print "Access Token = " . $oauth2->accessToken() . "\r\n";

# Get the full JSON response:
my $json = chilkat::CkJsonObject->new();
$json->Load( $oauth2->accessTokenResponse() );
$json->put_EmitCompact( 0 );

# print $json->emit() . "\r\n";

# Save the JSON to a file for future requests.
my $fac = chilkat::CkFileAccess->new();
$success = $fac->WriteEntireTextFile( "$sequence-token.json", $json->emit(), "utf-8", 0 );
if ($success != 1) {
    print $fac->lastErrorText();
    exit;
}

print "Finished.\n";
