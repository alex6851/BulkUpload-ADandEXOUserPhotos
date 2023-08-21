# BulkUpload-ADandEXOUserPhotos
This is for bulk uploading user photos to AD and EXO.

You can use this script to bulk upload user photos to users in AD and Exchange Online.

This script uploads to the users AD account and the Exchange Mailbox. If you have migrated everyone's mailbox to Exchange Online then you don't need to upload to users AD account. (Photos stored in the users mailbox are higher quality than the photos stored in AD) We only uploaded to both because we had custom built programs that used the photo from the AD user account.


## Adapt script for your usage

You will need to modify these lines:

line 328 for the SMTP server information.

line 353-385 You will need to change where your getting your photos from. I was getting them from a sharepoint share.


## Usage

.\BulkUpload-ADandEXOUserPhotos_2.1
