package exchangelib

import (
	"encoding/xml"
	"time"
)

// https://msdn.microsoft.com/en-us/library/office/aa563009(v=exchg.140).aspx

type CreateItem struct {
	XMLName            struct{}          `xml:"m:CreateItem"`
	MessageDisposition string            `xml:"MessageDisposition,attr"`
	SavedItemFolderId  SavedItemFolderId `xml:"m:SavedItemFolderId"`
	Items              Messages          `xml:"m:Items"`
}

type Messages struct {
	Message []Message `xml:"t:Message"`
}

type SavedItemFolderId struct {
	DistinguishedFolderId DistinguishedFolderId `xml:"t:DistinguishedFolderId"`
}

type DistinguishedFolderId struct {
	Id string `xml:"Id,attr"`
}

type Message struct {
	ItemClass                  string      `xml:"t:ItemClass"`
	Subject                    string      `xml:"t:Subject"`
	Body                       Body        `xml:"t:Body"`
	Attachments                Attachments `xml:"t:Attachments"`
	Sender                     OneMailbox  `xml:"t:Sender"`
	ToRecipients               XMailbox    `xml:"t:ToRecipients"`
	CcRecipients               XMailbox    `xml:"t:CcRecipients"`
	BccRecipients              XMailbox    `xml:"t:BccRecipients"`
	IsReadReceiptRequested     bool        `xml:"t:IsReadReceiptRequested"`     // whether a read receipt is requested for the e-mail message.
	IsDeliveryReceiptRequested bool        `xml:"t:IsDeliveryReceiptRequested"` // whether a delivery receipt is requested for the e-mail message.
}

type Attachments struct {
	FileAttachment []FileAttachment `xml:"t:FileAttachment"`
}

type Body struct {
	BodyType string `xml:"BodyType,attr"` // https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/body#bodytype
	Body     []byte `xml:",chardata"`
}

// ContentTypes refer to https://developer.mozilla.org/en-US/docs/Web/HTTP/Basics_of_HTTP/MIME_types/Common_types
type FileAttachment struct {
	Name             string    `xml:"t:Name"`
	ContentId        string    `xml:"t:ContentId"`
	ContentType      string    `xml:"t:ContentType"`     // default: "application/octet-stream"
	ContentLocation  string    `xml:"t:ContentLocation"` // the location of the content of an attachment
	Size             int32     `xml:"t:Size"`
	LastModifiedTime time.Time `xml:"t:LastModifiedTime"`
	IsInline         bool      `xml:"t:IsInline"`
	Content          string    `xml:"t:Content"`
}

type OneMailbox struct {
	Mailbox Mailbox `xml:"t:Mailbox"`
}

type XMailbox struct {
	Mailbox []Mailbox `xml:"t:Mailbox"`
}

type Mailbox struct {
	EmailAddress string `xml:"t:EmailAddress"`
	RoutingType  string `xml:"t:RoutingType"`
	MailboxType  string `xml:"t:MailboxType"`
}

func BuildTextEmail(
	from string,
	to []string,
	cc []string,
	bcc []string,
	subject string,
	body []byte,
	attachments []FileAttachment,
) ([]byte, error) {
	createItem := new(CreateItem)
	createItem.MessageDisposition = "SendAndSaveCopy"
	createItem.SavedItemFolderId.DistinguishedFolderId.Id = "sentitems"
	message := new(Message)
	message.ItemClass = "IPM.Note"
	message.Subject = subject
	message.Body.BodyType = "HTML"
	message.Body.Body = body
	message.Sender.Mailbox.EmailAddress = from
	message.Sender.Mailbox.RoutingType = "SMTP"
	message.Sender.Mailbox.MailboxType = "Mailbox"
	mailboxTo := make([]Mailbox, len(to))
	ccMailbox := make([]Mailbox, len(cc))
	bccMailbox := make([]Mailbox, len(bcc))
	for i, addr := range to {
		mailboxTo[i].EmailAddress = addr
	}
	for i, addr := range cc {
		ccMailbox[i].EmailAddress = addr
	}
	for i, addr := range bcc {
		bccMailbox[i].EmailAddress = addr
	}
	message.IsReadReceiptRequested = false
	message.IsDeliveryReceiptRequested = false
	message.ToRecipients.Mailbox = append(message.ToRecipients.Mailbox, mailboxTo...)
	message.CcRecipients.Mailbox = append(message.CcRecipients.Mailbox, ccMailbox...)
	message.BccRecipients.Mailbox = append(message.BccRecipients.Mailbox, bccMailbox...)
	message.Attachments.FileAttachment = append(message.Attachments.FileAttachment, attachments...)
	createItem.Items.Message = append(createItem.Items.Message, *message)
	return xml.MarshalIndent(createItem, "", "  ")
}
