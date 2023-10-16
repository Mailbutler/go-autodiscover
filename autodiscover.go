package autodiscover

import (
	"bytes"
	"crypto/tls"
	"encoding/base64"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"net/http"
	"strconv"
	"strings"

	log "github.com/sirupsen/logrus"
	"golang.org/x/exp/slices"
)

type DiscoveredInfo struct {
	EWSUrl          string
	ExchangeVersion string
}

func basicAuth(username, password string) string {
	auth := username + ":" + password
	return base64.StdEncoding.EncodeToString([]byte(auth))
}

func redirectedUrl(emailAddress string) string {
	domain := strings.Split(emailAddress, "@")[1]

	transport := &http.Transport{TLSClientConfig: &tls.Config{InsecureSkipVerify: true}}
	httpClient := &http.Client{Transport: transport}

	req, err := http.NewRequest("GET", fmt.Sprintf("https://autodiscover.%s/autodiscover/autodiscover.xml", domain), nil)
	if err != nil {
		return ""
	}

	resp, err := httpClient.Do(req)
	if err != nil || resp.StatusCode != 302 {
		return ""
	}

	return resp.Header.Get("Location")
}

func discoveryUrls(emailAddress string) [3]string {
	domain := strings.Split(emailAddress, "@")[1]

	return [3]string{
		fmt.Sprintf("https://%s/autodiscover/autodiscover.xml", domain),
		fmt.Sprintf("https://autodiscover.%s/autodiscover/autodiscover.xml", domain),
		redirectedUrl(emailAddress),
	}
}

type AutoDiscoveryRequest struct {
	XMLName                  xml.Name `xml:"http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006 Autodiscover"`
	EMailAddress             string   `xml:"Request>EMailAddress"`
	AcceptableResponseSchema string   `xml:"Request>AcceptableResponseSchema"`
}

func requestBody(emailAddress string) ([]byte, error) {
	requestBody := AutoDiscoveryRequest{
		EMailAddress:             emailAddress,
		AcceptableResponseSchema: "http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a",
	}
	return xml.Marshal(&requestBody)
}

type AutoDiscoveryResponse struct {
	XMLName   xml.Name                `xml:"http://schemas.microsoft.com/exchange/autodiscover/responseschema/2006 Autodiscover"`
	Protocols []AutoDiscoveryProtocol `xml:"Response>Account>Protocol"`
}

type AutoDiscoveryProtocol struct {
	XMLName       xml.Name `xml:"Protocol"`
	Type          string   `xml:"Type"`
	ServerVersion string   `xml:"ServerVersion"`
	EwsUrl        string   `xml:"EwsUrl"`
}

func exchangeVersion(rawHex string) (string, error) {
	i, err := strconv.ParseUint(rawHex, 16, 32)
	if err != nil {
		return "", err
	}

	binString := fmt.Sprintf("%032b", i)

	major, _ := strconv.ParseInt(binString[4:10], 2, 32)
	minor, _ := strconv.ParseInt(binString[10:16], 2, 32)
	build, _ := strconv.ParseInt(binString[17:32], 2, 32)

	switch major {
	case 8:
		switch minor {
		case 0:
			return "Exchange2007", nil
		case 1:
			return "Exchange2007_SP1", nil
		case 2:
			return "Exchange2007_SP1", nil
		case 3:
			return "Exchange2007_SP1", nil
		default:
			return "", errors.New("unknown minor version")
		}
	case 14:
		switch minor {
		case 0:
			return "Exchange2010", nil
		case 1:
			return "Exchange2010_SP1", nil
		case 2:
			return "Exchange2010_SP2", nil
		case 3:
			return "Exchange2010_SP2", nil
		default:
			return "", errors.New("unknown minor version")
		}
	case 15:
		switch minor {
		case 0:
			if build >= 847 {
				return "Exchange2013_SP1", nil // Minor builds starting from 847 are Exchange2013_SP1
			} else {
				return "Exchange2013", nil
			}
		case 1:
			return "Exchange2016", nil
		case 2:
			return "Exchange2019", nil
		case 20:
			return "Exchange2016", nil // This is Office365
		default:
			return "", errors.New("unknown minor version")
		}
	default:
		return "", errors.New("unknown major version")
	}
}

func parseResponse(body []byte) (DiscoveredInfo, error) {
	var response AutoDiscoveryResponse
	err := xml.Unmarshal(body, &response)
	if err != nil {
		return DiscoveredInfo{}, err
	}

	protocols := response.Protocols

	exchangeProtocolIndex := slices.IndexFunc(protocols, func(p AutoDiscoveryProtocol) bool { return p.Type == "EXCH" })
	if exchangeProtocolIndex < 0 {
		return DiscoveredInfo{}, errors.New("failed to find Exchange protocol in response")
	}
	exchangeProtocol := protocols[exchangeProtocolIndex]
	ewsUrl := exchangeProtocol.EwsUrl
	rawVersion := exchangeProtocol.ServerVersion

	if ewsUrl == "" {
		expressProtocolIndex := slices.IndexFunc(protocols, func(p AutoDiscoveryProtocol) bool { return p.Type == "EXPR" })
		if expressProtocolIndex < 0 {
			return DiscoveredInfo{}, errors.New("failed to find Express protocol in response")
		}
		expressProtocol := protocols[expressProtocolIndex]

		ewsUrl = expressProtocol.EwsUrl
	}

	exchangeVersion, err := exchangeVersion(rawVersion)
	if err != nil {
		return DiscoveredInfo{}, err
	}

	return DiscoveredInfo{EWSUrl: ewsUrl, ExchangeVersion: exchangeVersion}, nil
}

func Discover(emailAddress string, password string) (DiscoveredInfo, error) {
	transport := &http.Transport{TLSClientConfig: &tls.Config{InsecureSkipVerify: true}}
	httpClient := &http.Client{Transport: transport}

	body, err := requestBody(emailAddress)
	if err != nil {
		return DiscoveredInfo{}, err
	}

	for _, url := range discoveryUrls(emailAddress) {
		if url == "" {
			continue
		}

		req, err := http.NewRequest("GET", url, bytes.NewBuffer(body))
		if err != nil {
			log.WithError(err).Warn("failed to create GET request for " + url)
			continue
		}

		req.Header.Add("Authorization", "Basic "+basicAuth(emailAddress, password))
		resp, err := httpClient.Do(req)
		if err != nil {
			log.WithError(err).Warn("failed to receive response from " + url)
			continue
		}

		if resp.StatusCode != 200 {
			continue
		}

		body, readErr := io.ReadAll(resp.Body)
		if readErr != nil {
			log.WithError(err).Warn("failed to read response from " + url)
			continue
		}

		info, parseErr := parseResponse(body)
		if parseErr != nil {
			log.WithError(err).Warn("failed to parse response from " + url)
			continue
		}

		return info, nil
	}

	return DiscoveredInfo{}, errors.New("failed to fetch discovery information")
}
