# クレジットカード利用通知ボット

特定の銀行(例: 三井住友カード)からのクレジットカード利用通知メールをGmailから監視し、メール本文から正規表現を使って日付、場所、金額などの取引詳細を抽出します。抽出したデータはGoogle SpreadsheetにMonth単位のシートで記録されます。毎日の終わりに、ボットは指定されたLINEチャットに1日の利用金額と月の合計金額を要約したメッセージを送信します。

## 主な機能

- Gmail APIを使ってクレジットカード利用通知メールを検索・取得
- 正規表現を使ってメール本文から取引詳細を解析
- 取引データをMonth単位でGoogle Spreadsheetに保存
- シート内の取引を日付順に自動でソート
- LINE Messaging APIを使って指定のLINEチャットに日次レポートを送信
- 処理済みのメッセージIDを追跡して重複を回避

## 使用技術

- Google Apps Script
- Gmail API
- Google Sheets API
- LINE Messaging API
- 正規表現

このボットはクレジットカード利用状況の確認を簡素化し、手動でのデータ入力の必要がなくなり、LINEを通じて便利に日次サマリーを提供します。

# credit-card-usage-notification-bot

A bot that monitors Gmail for credit card usage notifications from a specific bank (e.g., Sumitomo Mitsui Card). It extracts transaction details like date, location, and amount from the email body using regular expressions. The extracted data is then recorded in a Google Spreadsheet, with one sheet per month. At the end of each day, the bot sends a message to a specified LINE chat, summarizing the daily and monthly credit card usage totals.

## Key Features

- Utilizes Gmail API to search for and fetch credit card usage notification emails
- Employs regular expressions to parse transaction details from email body
- Stores transaction data in a Google Spreadsheet, organized by month
- Automatically sorts transactions in chronological order within each sheet
- Sends daily reports to a specified LINE chat using LINE Messaging API
- Avoids duplicates by tracking processed email message IDs

## Tech Stack

- Google Apps Script
- Gmail API
- Google Sheets API
- LINE Messaging API
- Regular Expressions

This bot simplifies the process of monitoring credit card usage, eliminating the need for manual data entry and providing daily summaries conveniently through LINE.
