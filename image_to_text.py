import google.generativeai as genai
from PIL import Image
import os
import json
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
import datetime
import re

# .envファイルから環境変数を読み込む
load_dotenv()

# Gemini APIの設定
GOOGLE_API_KEY = os.getenv('GOOGLE_API_KEY')
genai.configure(api_key=GOOGLE_API_KEY)

# Google Sheets APIのスコープ
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def get_google_sheets_service():
    """
    Google Sheets APIのサービスアカウント認証を行い、サービスオブジェクトを返す
    
    Returns:
        googleapiclient.discovery.Resource: Google Sheets APIのサービスオブジェクト
    """
    # サービスアカウントファイルのパスを環境変数から取得
    service_account_file = os.getenv('SERVICE_ACCOUNT_FILE')
    if not service_account_file:
        raise ValueError("SERVICE_ACCOUNT_FILE environment variable is not set")
    
    # サービスアカウントの認証情報を読み込む
    credentials = service_account.Credentials.from_service_account_file(
        service_account_file, scopes=SCOPES)
    
    return build('sheets', 'v4', credentials=credentials)

def convert_to_table(text):
    """
    テキストを表形式に変換する関数
    
    Args:
        text (str): 変換するテキスト
    
    Returns:
        list: 表形式のデータ（各行がリストのリスト）
    """
    # 行ごとに分割
    lines = text.strip().split('\n')
    
    # 各行をタブまたはカンマで分割
    table_data = []
    for line in lines:
        # タブまたはカンマで分割
        row = re.split(r'[\t,]', line)
        # 空の要素を削除
        row = [item.strip() for item in row if item.strip()]
        if row:  # 空の行は追加しない
            table_data.append(row)
    
    return table_data

def append_to_sheet(spreadsheet_id, range_name, values):
    """
    スプレッドシートにデータを追加する
    
    Args:
        spreadsheet_id (str): スプレッドシートのID
        range_name (str): データを追加する範囲
        values (list): 追加するデータのリスト
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

def extract_text_from_image(image_path):
    """
    画像からテキストを抽出する関数
    
    Args:
        image_path (str): 画像ファイルのパス
    
    Returns:
        str: 抽出されたテキスト
    """
    try:
        # 画像を開く
        image = Image.open(image_path)
        
        # Geminiモデルの初期化（最新のモデルを使用）
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        # 画像からテキストを抽出
        response = model.generate_content(["この画像に含まれるテキストを抽出して表形式で出力してください。各列はタブまたはカンマで区切ってください。", image])
        
        return response.text
    
    except Exception as e:
        return f"エラーが発生しました: {str(e)}"

if __name__ == "__main__":
    # 使用例
    image_path = "sample.png"  # 画像ファイルのパスを指定
    extracted_text = extract_text_from_image(image_path)
    print("抽出されたテキスト:")
    print(extracted_text)
    
    # テキストを表形式に変換
    table_data = convert_to_table(extracted_text)
    
    # スプレッドシートにデータを追加
    SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')  # .envファイルからスプレッドシートIDを取得
    RANGE_NAME = 'シート1!A1'  # データを追加する開始位置
    
    # 現在の日時を取得
    current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    # スプレッドシートに追加するデータ
    # 最初の行に日時と画像パスを追加
    values = [[current_time, image_path]]
    # 表データを追加
    values.extend(table_data)
    
    try:
        append_to_sheet(SPREADSHEET_ID, RANGE_NAME, values)
        print("スプレッドシートにデータを追加しました")
    except Exception as e:
        print(f"スプレッドシートへの追加中にエラーが発生しました: {str(e)}") 