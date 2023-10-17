package autodiscover

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestDiscover(t *testing.T) {
	info, _ := Discover("testmailbutler@wp13554064.server-he.de", "geheim.MB44")

	assert.Equal(t, "https://mail.hexchange.de/EWS/Exchange.asmx", info.EWSUrl)
	assert.Equal(t, "Exchange2019", info.ExchangeVersion)
}

func TestTestExchangeVersion(t *testing.T) {
	versionA, _ := exchangeVersion("738180DA")
	assert.Equal(t, "Exchange2010_SP1", versionA)

	versionB, _ := exchangeVersion("73C1840A")
	assert.Equal(t, "Exchange2016", versionB)
}

func TestRequestBody(t *testing.T) {
	body, _ := requestBody("somebody@gmail.com")

	expectedBody := `<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006"><Request><EMailAddress>somebody@gmail.com</EMailAddress><AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema></Request></Autodiscover>`

	assert.Equal(t, expectedBody, string(body))
}

func TestParseResponse(t *testing.T) {
	responseBody := `<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/responseschema/2006">
  <Response xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a">
    <User>
      <DisplayName>User1, vornme</DisplayName>
      <LegacyDN>/o=orgname/ou=admingroup/cn=Recipients/cn=User1</LegacyDN>
      <AutoDiscoverSMTPAddress>User1@msxfaq.net</AutoDiscoverSMTPAddress>
      <DeploymentId>12345678-1234-1234-1234-123456789012</DeploymentId>
    </User>
    <Account>
      <AccountType>email</AccountType>
      <Action>settings</Action>
      <MicrosoftOnline>False</MicrosoftOnline>
      <ConsumerMailbox>False</ConsumerMailbox>
	  <Protocol>
        <Type>EXCH</Type>
        <Server>293f8bce-0287-458f-a98c-a70f927e00fd@krone.de</Server>
        <ServerDN>/o=IT-P-S/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Configuration/cn=Servers/cn=293f8bce-0287-458f-a98c-a70f927e00fd@krone.de</ServerDN>
        <ServerVersion>73C1840A</ServerVersion>
        <MdbDN>/o=IT-P-S/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Configuration/cn=Servers/cn=293f8bce-0287-458f-a98c-a70f927e00fd@krone.de/cn=Microsoft Private MDB</MdbDN>
        <PublicFolderServer>outlook.msxfaq.net</PublicFolderServer>
        <AD>dc01.msxfaq.net</AD>
        <ASUrl>https://outlook.msxfaq.net/EWS/Exchange.asmx</ASUrl>
        <EmwsUrl>https://outlook.msxfaq.net/EWS/Exchange.asmx</EmwsUrl>
        <EcpUrl>https://outlook.msxfaq.net/owa/</EcpUrl>
        <EcpUrl-um>?path=/options/callanswering</EcpUrl-um>
        <EcpUrl-aggr>?path=/options/connectedaccounts</EcpUrl-aggr>
        <EcpUrl-mt>options/ecp/PersonalSettings/DeliveryReport.aspx?rfr=olk&amp;exsvurl=1&amp;IsOWA=&lt;IsOWA&gt;&amp;MsgID=&lt;MsgID&gt;&amp;Mbx=&lt;Mbx&gt;&amp;realm=krone.de</EcpUrl-mt>
        <EcpUrl-ret>?path=/options/retentionpolicies</EcpUrl-ret>
        <EcpUrl-photo>?path=/options/myaccount/action/photo</EcpUrl-photo>
        <EcpUrl-extinstall>?path=/options/manageapps</EcpUrl-extinstall>
        <OOFUrl>https://outlook.msxfaq.net/EWS/Exchange.asmx</OOFUrl>
        <UMUrl>https://outlook.msxfaq.net/EWS/UM2007Legacy.asmx</UMUrl>
        <OABUrl>https://outlook.msxfaq.net/oab/12345678-1234-1234-1234-123456789012/</OABUrl>
        <ServerExclusiveConnect>off</ServerExclusiveConnect>
      </Protocol>
	  <Protocol>
   		<Type>EXPR</Type>
   		<Server>rpc.msxfaq.net</Server>
   		<SSL>On</SSL>
   		<AuthPackage>Ntlm</AuthPackage>
   		<ServerExclusiveConnect>on</ServerExclusiveConnect>
   		<GroupingInformation>DE-Paderborn</GroupingInformation>
		<EwsUrl>https://outlook.msxfaq.net/EWS/Exchange.asmx</EwsUrl>
	  </Protocol>
    </Account>
  </Response>
</Autodiscover>
	`

	info, _ := parseResponse([]byte(responseBody))

	assert.Equal(t, "https://outlook.msxfaq.net/EWS/Exchange.asmx", info.EWSUrl)
	assert.Equal(t, "Exchange2016", info.ExchangeVersion)
}
