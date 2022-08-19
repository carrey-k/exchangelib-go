# exchangelib-go
Go client for Microsoft Exchange Web Services (EWS)

## Installation

### Setting Up Go
To install Go, visit [this link](https://golang.org/dl/).

### Installing Module
`go get -u github.com/kangchengkun/exchangelib-go`

## Usage
Before using this Go module, you will need to fetch a access token from microsoft online by using https://github.com/kangchengkun/mso-token.


```
import github.com/kangchengkun/exchangelib-go

exchangelib.Sender = "your-mail-box"
exchangelib.AccessToken = "your-ews-token"

// Change the default exchange web service endpoint
exchangelib.ExchangeServerAddr = "your-exchange-server"


// Send an email
response, err := exchangelib.SendMail(to, cc, bcc, subject, body, attachments)
if err != nil {
    fmt.Println("SendMail failed")
}
```

## Contribution

Follow the [Guide](https://go.dev/blog/publishing-go-modules) to publish new versions

```
...
git add .
git commit -m "new updates"

$ git tag vx.x.x
$ git push origin vx.x.x
```