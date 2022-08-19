package exchangelib

import (
	"bytes"
	"crypto/tls"
	"errors"
	"fmt"
	"net/http"
	"regexp"
	"strings"
	"time"

	httpNtlm "github.com/vadimi/go-http-ntlm"
)

var (
	Sender             string // mail or domain\account format
	AccessToken        string // the access token for exchange web services
	ExchangeServerAddr string = "https://outlook.office365.com/EWS/Exchange.asmx"
)

var soapHeader = `<?xml version="1.0" encoding="utf-8" ?>
<s:Envelope
	xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
	xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
	xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
	<s:Header>
		<t:RequestServerVersion Version="Exchange2016"/>
		<t:ExchangeImpersonation>
			<t:ConnectingSID>
				<t:PrimarySmtpAddress>[account-placeholder]</t:PrimarySmtpAddress>
			</t:ConnectingSID>
		</t:ExchangeImpersonation>
		<t:TimeZoneContext>
			<t:TimeZoneDefinition Id="China Standard Time"/>
		</t:TimeZoneContext>
	</s:Header>
	<s:Body> 
`

func SendMail(
	to []string,
	cc []string,
	bcc []string,
	topic string,
	content string,
	attachments []FileAttachment,
) (*http.Response, error) {
	if Sender == "" {
		return nil, errors.New("no valid sender provided")
	}

	b, err := BuildTextEmail(Sender, to, cc, bcc, topic, []byte(content), attachments)
	if err != nil {
		return nil, errors.New(fmt.Sprintf("build text email failed: %s", err))
	}

	return Issue(ExchangeServerAddr, b)
}

func Issue(ewsAddr string, body []byte) (*http.Response, error) {
	if Sender == "" {
		return nil, errors.New("empty user name, please provide valid email or format with domain\\account")
	}
	if ewsAddr == "" {
		return nil, errors.New("empty ews address, please provide valid server address")
	}

	if AccessToken == "" {
		return nil, errors.New("empty ews access token, please provide valid access token")
	}
	header := strings.ReplaceAll(soapHeader, "[account-placeholder]", Sender)
	b2 := []byte(header)
	b2 = append(b2, body...)
	b2 = append(b2, "\n  </s:Body>\n</s:Envelope>"...)
	req, err := http.NewRequest("POST", ewsAddr, bytes.NewReader(b2))
	if err != nil {
		return nil, errors.New(fmt.Sprintf("create request failed: %s", err))
	}

	var client *http.Client
	re := regexp.MustCompile("^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$")
	isMail := re.MatchString(Sender)
	bearer := "Bearer " + AccessToken
	if !isMail {
		// use domain
		l := strings.Split(Sender, "\\")
		if len(l) < 2 {
			return nil, errors.New("wrong format of username, not email or format with domain\\account")
		}

		domain := l[0]
		account := l[1]
		client = &http.Client{
			Transport: &httpNtlm.NtlmTransport{
				Domain:          domain,
				User:            account,
				TLSClientConfig: &tls.Config{InsecureSkipVerify: true},
			},
		}
	} else {
		client = &http.Client{
			Transport: &http.Transport{TLSClientConfig: &tls.Config{InsecureSkipVerify: true}},
		}
	}
	req.Header.Add("Authorization", bearer)
	req.Header.Add("X-AnchorMailbox", Sender)
	req.Header.Set("Content-Type", "text/xml; charset=utf-8")
	second := time.Second
	client.Timeout = 120 * second // request timeout within 2 minutes.
	client.CheckRedirect = func(req *http.Request, via []*http.Request) error { return http.ErrUseLastResponse }
	fmt.Println("Start to send http request to EWS server...")
	return client.Do(req)
}
