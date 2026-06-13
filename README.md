# 🍌 Nanobanana Editor (ナノバナナ エディター)

Nanobanana Editor は、インフォグラフィック画像（PNG/JPG）や PDF ファイルを解析し、**「文字の自動消去・背景復元（インペインティング）」** と **「編集可能な PowerPoint テキストオブジェクトへの変換」** を同時に行うWebアプリケーションです。

生成された PowerPoint（`.pptx`）ファイルは、Microsoft PowerPoint や Google スライドにインポートして、テキストや色、サイズを 100% 自由に再編集することができます。

---

## 1. ソフトの仕様 (Software Specifications)

### 💻 開発スタック
* **フロントエンド:** Streamlit (カスタム CSS によるモダンで美しい Glassmorphism デザイン)
* **バックエンド:** Python 3.10+
* **OCR・座標検出エンジン:** Google Cloud Vision API
* **画像処理 (文字消去 & 背景復元):** OpenCV (Telea Inpainting ロジック)
* **PDF 解析:** PyMuPDF (`fitz` / システム依存バイナリ `poppler` 不要の純 Python 実装)
* **Officeファイル生成:** `python-pptx`
* **プレビュー文字描画:** Pillow (`PIL.ImageDraw` / macOS 向けに NFD パスによるヒラギノフォント対応)

### 🌟 主な機能要件
1. **高精度な OCR と位置検出**: アップロードされたファイルから文字の位置（Bounding Box）とテキストを抽出・シリアライズします。
2. **自動背景補完 (Inpainting)**: 検出された文字の周囲数ピクセルを含めてマスクし、OpenCV を用いて違和感なく文字を消去し、背景のみの画像を自動生成します。
3. **リアルタイム仕上がりプレビュー (Live Preview)**: 画面左側のプレビューにて、消去後の背景にユーザーが編集したテキスト（フォント・サイズ・カラー）をリアルタイムで重ね書きして確認できます。
4. **自動フォントスケーリング**: テキスト長が増加した際、元の配置枠（$W_{box}$）をテキスト幅（$W_{text}$）が超えてレイアウトが崩れないよう、以下の数式を用いてフォントサイズを自動縮小します。
   $$S_{new} = S_{init} \times \min\left(1.0, \frac{W_{box}}{W_{text}}\right)$$
5. **手動微調整 UI**: 各ブロックごとにフォントサイズを 1pt 単位で微調整できるスライダー、カラーピッカー、およびフォントファミリー代替案（Gothic、Mincho、Round、Design、およびカスタム入力）を用意しています。
6. **マルチ出力対応**: `.pptx` スライドのほか、編集後の画像を `.png`、`.jpg` 形式で直接ダウンロード可能です。

---

## 2. 環境設定方法 (Environment Setup)

本アプリケーションでカスタム画像/PDFのOCR解析を行うには、**Google Cloud Vision API** の有効化および認証設定が必要です。以下のいずれかの方法で設定を行います。

### 🔑 方法A：サービスアカウントの JSON 鍵ファイルをアップロードする (簡単)
1. [Google Cloud Console](https://console.cloud.google.com/) で **Cloud Vision API** を有効にします。
2. サービスアカウントを作成し、**JSON 形式の秘密鍵キー**をダウンロードします。
3. アプリケーション（Streamlit）起動後、左側のサイドバーを展開し、「デモモードで試す」を OFF にして、ダウンロードした JSON ファイルをアップロード欄にドラッグ＆ドロップします。

### 🛡️ 方法B：Workload Identity / 環境認証情報 (ADC) を連携する (セキュア)
サーバー環境（Cloud Run、GKE など）の Workload Identity や、ローカル PC のアプリケーションデフォルト資格情報（ADC）を直接使用できます。

1. **gcloud CLI のインストール** (macOS の場合は `brew install --cask google-cloud-sdk` など)。
2. ターミナルで初期設定と認証を完了します：
   ```bash
   # 1. ログイン設定 (ブラウザで認証)
   gcloud init

   # 2. アプリケーションデフォルト認証情報 (ADC) の作成
   gcloud auth application-default login

   # 3. 使用する GCP プロジェクト (利用クォータ) の設定
   gcloud auth application-default set-quota-project <YOUR_PROJECT_ID>

   # 4. 指定したプロジェクトで Vision API を有効化
   gcloud services enable vision.googleapis.com
   ```
3. 上記設定が完了している環境では、**JSON キーファイルのアップロードなしで自動的に API 認証が通ります**。

---

## 3. インストール方法 (Installation)

### ① リポジトリのクローン
```bash
git clone https://github.com/MT-Git01/Nanobanana_Editor.git
cd Nanobanana_Editor
```

### ② 依存パッケージのインストール
Python 3.10 以上がインストールされている環境で、以下を実行します：
```bash
pip install streamlit opencv-python python-pptx google-cloud-vision PyMuPDF pillow
```

---

## 4. ソフトの利用方法 (Usage)

### ① サーバーの起動
プロジェクトディレクトリ内で、以下のコマンドを実行して Streamlit ローカル開発サーバーを起動します：
```bash
streamlit run app.py
```
起動すると、自動的にブラウザで `http://localhost:8502`（または `8501`）が開きます。

### ② デモモードでの体験
* 起動時、サイドバーの「デモモードで試す (Mock OCR)」がオンになっています。
* サンプルのインフォグラフィック画像を用いて、文字消去、リアルタイム編集、PPTX・画像の書き出し動作をすぐに体験できます。

### ③ 実ファイルでの実行
1. サイドバーの「デモモードで試す」を OFF にします。
2. 上記「2. 環境設定方法」に基づき、JSON キーをアップロードするか、PC の認証情報を有効にします。
3. 解析したい画像（PNG, JPG, JPEG）または PDF ファイルをドラッグ＆ドロップでアップロードします。
4. OCR解析と背景消去（Inpainting）が自動的に実行されます。

### ④ 編集と書き出し
1. **プレビュー表示モード**: 
   * `編集後のリアルタイム仕上がり (Live Preview)`: 編集結果を画像上でリアルタイムにプレビュー確認します。
   * `検出枠のハイライト（オリジナル）`: 検出された黄色の領域枠とブロックIDを確認します。
2. **テキストブロックの編集**:
   * 右側パネルから編集したいブロックを選択し、テキスト内容の書き換え、代替フォントの設定、サイズの微調整、テキストカラーの変更を行います。
3. **ダウンロード**:
   * パネル下部のダウンロードボタンから、`.pptx` ファイル（PowerPoint / Googleスライドにインポートして再編集可能）、`.png` 画像、`.jpg` 画像をそれぞれワンクリックでローカルディスクに保存します。

---

## 5. GitHub リポジトリ (GitHub Repository)

コードの管理、コントリビューション、最新バージョンの取得は以下のリポジトリで行っています：

🔗 [https://github.com/MT-Git01/Nanobanana_Editor.git](https://github.com/MT-Git01/Nanobanana_Editor.git)
