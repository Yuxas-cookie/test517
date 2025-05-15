from flask import Flask, request, abort
from linebot.v3 import WebhookHandler
from linebot.v3.messaging import (
    Configuration,
    ApiClient,
    MessagingApi,
    ReplyMessageRequest,
    TextMessage
)
from linebot.v3.exceptions import InvalidSignatureError
from linebot.v3.webhooks import MessageEvent, ImageMessageContent
import os
from dotenv import load_dotenv
from datetime import datetime
import logging
import requests
import json
import google.generativeai as genai
from PIL import Image
import io
from google.oauth2 import service_account
from googleapiclient.discovery import build
import re

# ログ設定
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# .envファイルから環境変数を読み込む
load_dotenv()

app = Flask(__name__)

# LINE Messaging APIの設定
configuration = Configuration(access_token=os.getenv('LINE_CHANNEL_ACCESS_TOKEN'))
handler = WebhookHandler(os.getenv('LINE_CHANNEL_SECRET'))

# Gemini APIの設定
GOOGLE_API_KEY = os.getenv('GOOGLE_API_KEY')
genai.configure(api_key=GOOGLE_API_KEY)

# Google Sheets APIのスコープ
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# 環境変数の確認
logger.info(f"Channel Access Token: {os.getenv('LINE_CHANNEL_ACCESS_TOKEN')[:10]}...")
logger.info(f"Channel Secret: {os.getenv('LINE_CHANNEL_SECRET')[:10]}...")
logger.info(f"Google API Key: {os.getenv('GOOGLE_API_KEY')[:10]}...")
logger.info(f"Spreadsheet ID: {os.getenv('SPREADSHEET_ID')[:10]}...")
logger.info(f"Service Account File: {os.getenv('SERVICE_ACCOUNT_FILE')}")

# 画像を保存するディレクトリ
SAVE_DIR = 'images'
if not os.path.exists(SAVE_DIR):
    os.makedirs(SAVE_DIR)

# 保存済みメッセージIDを記録するファイル
SAVED_MESSAGES_FILE = 'saved_messages.json'

def get_google_sheets_service():
    """
    Google Sheets APIのサービスアカウント認証を行い、サービスオブジェクトを返す
    """
    # サービスアカウントファイルのパスを環境変数から取得
    service_account_file = os.getenv('SERVICE_ACCOUNT_FILE')
    if not service_account_file:
        raise ValueError("SERVICE_ACCOUNT_FILE environment variable is not set")
    
    # サービスアカウントの認証情報を読み込む
    credentials = service_account.Credentials.from_service_account_file(
        service_account_file, scopes=SCOPES)
    print(credentials)
    return build('sheets', 'v4', credentials=credentials)

def convert_to_table(text):
    """
    テキストを表形式に変換する関数
    """
    lines = text.strip().split('\n')
    table_data = []
    for line in lines:
        row = re.split(r'[\t,]', line)
        row = [item.strip() for item in row if item.strip()]
        if row:
            table_data.append(row)
    return table_data

def append_to_sheet(spreadsheet_id, range_name, values):
    """
    スプレッドシートにデータを追加する
    """
    service = get_google_sheets_service()
    body = {
        'values': values
    }
    result = service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption='RAW',
        body=body
    ).execute()
    return result

def extract_text_from_image(image_data):
    """
    画像からテキストを抽出する関数
    """
    try:
        # 画像データからPIL Imageオブジェクトを作成
        image = Image.open(io.BytesIO(image_data))
        
        # Geminiモデルの初期化
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        # 画像からテキストを抽出
        response = model.generate_content(["この画像に含まれるテキストを抽出して表形式で出力してください。各列はタブまたはカンマで区切ってください。", image])
        
        return response.text
    
    except Exception as e:
        logger.error(f"Error extracting text from image: {str(e)}")
        return None

# 保存済みメッセージIDを読み込む
def load_saved_messages():
    if os.path.exists(SAVED_MESSAGES_FILE):
        with open(SAVED_MESSAGES_FILE, 'r') as f:
            return json.load(f)
    return []

# 保存済みメッセージIDを保存する
def save_message_id(message_id):
    saved_messages = load_saved_messages()
    if message_id not in saved_messages:
        saved_messages.append(message_id)
        with open(SAVED_MESSAGES_FILE, 'w') as f:
            json.dump(saved_messages, f)

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers.get('X-Line-Signature')
    if not signature:
        logger.error("No signature header found")
        abort(400)

    body = request.get_data(as_text=True)
    logger.info("Request body: " + body)
    logger.info("Signature: " + signature)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError as e:
        logger.error(f"Invalid signature: {str(e)}")
        abort(400)
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        abort(500)

    return 'OK'

@handler.add(MessageEvent, message=ImageMessageContent)
def handle_image_message(event):
    try:
        message_id = event.message.id
        logger.info(f"Received image message: {message_id}")
        
        # 既に保存済みのメッセージかチェック
        saved_messages = load_saved_messages()
        if message_id in saved_messages:
            logger.info(f"Image already processed: {message_id}")
            return
        
        # 画像のコンテンツを取得
        headers = {
            'Authorization': f'Bearer {os.getenv("LINE_CHANNEL_ACCESS_TOKEN")}'
        }
        url = f'https://api-data.line.me/v2/bot/message/{message_id}/content'
        
        # 画像をダウンロード
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            # 画像データを取得
            image_data = response.content
            
            # 画像からテキストを抽出
            extracted_text = extract_text_from_image(image_data)
            
            if extracted_text:
                # テキストを表形式に変換
                table_data = convert_to_table(extracted_text)
                
                # スプレッドシートにデータを追加
                SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')
                RANGE_NAME = 'シート1!A1'
                
                # 現在の日時を取得
                current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                
                # スプレッドシートに追加するデータ
                values = [[current_time, f"LINE Message ID: {message_id}"]]
                values.extend(table_data)
                
                try:
                    append_to_sheet(SPREADSHEET_ID, RANGE_NAME, values)
                    logger.info("Data added to spreadsheet successfully")
                    
                    # ユーザーに成功を通知
                    with ApiClient(configuration) as api_client:
                        messaging_api = MessagingApi(api_client)
                        messaging_api.reply_message(
                            ReplyMessageRequest(
                                reply_token=event.reply_token,
                                messages=[TextMessage(text='画像からテキストを抽出し、スプレッドシートに保存しました。')]
                            )
                        )
                except Exception as e:
                    logger.error(f"Error adding to spreadsheet: {str(e)}")
                    raise
            else:
                # テキスト抽出に失敗した場合
                with ApiClient(configuration) as api_client:
                    messaging_api = MessagingApi(api_client)
                    messaging_api.reply_message(
                        ReplyMessageRequest(
                            reply_token=event.reply_token,
                            messages=[TextMessage(text='テキストの抽出に失敗しました。')]
                        )
                    )
            
            # メッセージIDを保存済みとして記録
            save_message_id(message_id)
            
        else:
            logger.error(f"Failed to download image. Status code: {response.status_code}")
            raise Exception(f"Failed to download image. Status code: {response.status_code}")
            
    except Exception as e:
        logger.error(f"Error handling image message: {str(e)}")
        with ApiClient(configuration) as api_client:
            messaging_api = MessagingApi(api_client)
            messaging_api.reply_message(
                ReplyMessageRequest(
                    reply_token=event.reply_token,
                    messages=[TextMessage(text='処理中にエラーが発生しました。')]
                )
            )

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=8000, debug=True) 