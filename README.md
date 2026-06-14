# 🍌 Nanobanana Editor (ナノバナナ エディター)

Nanobanana Editor は、インフォグラフィック画像（PNG/JPG）や PDF ファイルを解析し、**「ユーザーが選択した文字の消去・背景復元（選択的インペインティング）」** と **「編集可能な PowerPoint テキストオブジェクトへの変換」** を同時に行うWebアプリケーションです。

インフォグラフィック全体のデザインを崩すことなく、テキストを修正したい箇所だけを選択して PowerPoint（`.pptx`）スライドや高解像度画像を生成できます。生成されたスライドは、Microsoft PowerPoint や Google スライドにインポートして、100% 自由に再編集可能です。

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
1. **選択的インペインティング (Selective Inpainting)**:
   * OCR が検出したすべての文字を無条件で消去するのではなく、**ユーザーが編集したいブロックのみを選択**して消去・背景復元を行います。
   * 編集対象に選ばれなかったテキストは、そのまま平坦な背景画像の一部として残るため、複雑なグラフィックや装飾文字のデザイン崩れを完全に防ぎます。
2. **高解像度の完全維持 (1K, 2K, 4K 対応)**:
   * 読み込んだ画像の解像度を縮小せずそのまま処理します。4K（3840x2160）などの高解像度ファイルでも、鮮明な画質のまま背景補完や書き出しが可能です。
3. **切り抜きプレビューと透過処理**:
   * 編集対象として選択されたテキスト領域を透過処理（アルファチャンネル 0）し、チェッカーボード（市松模様）背景を重ねて表示することで、どの部分がスライドのテキスト枠に置き換わるかを直感的に把握できます。
   * 選択した枠はコーラル（赤）、未選択の候補枠はイエロー（黄）でプレビュー上に明示されます。
4. **スライド比率の自動整合機能 (Auto Aspect Ratio)**:
   * スライドの比率に「元画像に合わせる (Auto)」オプションを搭載。元画像の縦横比に合わせて PowerPoint のスライドサイズを動的に決定するため、画像が途中で切れたり、縦横に引き伸ばされて歪むトラブルが起こりません。
5. **PowerPoint 向けフォント自動比例縮小 (Proportional Font Scaling)**:
   * ユーザーのテキスト編集による文字数溢れを防ぐ自動フォント縮小ロジックに加え、高解像度（4K等）のピクセル座標を PowerPoint のスライドポイント単位（pt）へ正確に換算し、文字サイズが枠からはみ出すのを防ぎます。
6. **位置・サイズのシフト微調整 (Position & Size Offsets)**:
   * 各テキストブロックごとに、X軸・Y軸方向への位置シフト（上下左右移動）および配置枠の幅の追加/削減（ピクセル単位）を個別設定できます。これにより、OCRの検出座標にズレがある場合でも、完璧な文字配置が可能です。
7. **ローカルPC「ダウンロード」フォルダへの直接保存**:
   * ブラウザのセキュリティ設定等によりダウンロードが妨げられるケースに備え、ワンクリックでローカルPCのシステム「ダウンロード」フォルダ（`~/Downloads`）へ直接ファイルを書き出す機能を搭載しています。

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

### ② 環境構築と依存パッケージのインストール (Miniconda + conda-forge 推奨)

企業内での商用利用など、ライセンスの観点から **Miniconda** および **conda-forge** チャンネルの使用を推奨します（Anaconda Commercial Edition のライセンス制限を受けず、商用利用無料です）。

1. **Miniconda のインストール**:
   お使いの OS に合わせて [Miniconda 公式ページ](https://docs.anaconda.com/miniconda/) から Miniconda をダウンロード・インストールしてください。

2. **仮想環境の作成と有効化**:
   Python 3.10 の環境を作成し、アクティベートします。
   ```bash
   conda create -n nanobanana-env python=3.10 -c conda-forge -y
   conda activate nanobanana-env
   ```

3. **依存パッケージのインストール**:
   conda-forge チャンネルから必要なパッケージを一括でインストールします。
   ```bash
   conda install -c conda-forge streamlit opencv python-pptx google-cloud-vision pymupdf pillow numpy -y
   ```

*(※ pip を利用したい場合は、同梱されている `requirements.txt` を使用して一括インストール可能です。)*
   ```bash
   pip install -r requirements.txt
   ```

---

## 4. ソフトの利用方法 (Usage)

### ① サーバーの起動
作成した仮想環境（Minicondaなど）をアクティベートし、プロジェクトディレクトリ内で以下のコマンドを実行して Streamlit ローカル開発サーバーを起動します：
```bash
conda activate nanobanana-env
streamlit run app.py
```
*(※ pip環境で構築した場合は、そのまま `streamlit run app.py` を実行してください。)*

起動すると、自動的にブラウザで `http://localhost:8502`（または `8501`）が開きます。

### ② デモモードでの体験
* 起動時、サイドバーの「デモモードで試す (Mock OCR)」がオンになっています。
* サンプルのインフォグラフィック画像を用いて、選択的インペインティング、リアルタイム編集、PPTX・画像の書き出し動作をすぐに体験できます。

### ③ 実ファイルでの実行
1. サイドバーの「デモモードで試す」を OFF にします。
2. 上記「2. 環境設定方法」に基づき、JSON キーをアップロードするか、PC の認証情報を有効にします。
3. 解析したい画像（PNG, JPG, JPEG）または PDF ファイルをドラッグ＆ドロップでアップロードします。
4. OCR 解析が自動的に実行され、検出されたテキストに黄色の枠線が引かれます。

### ④ テキスト選択と編集
1. **左カラム (切り抜きプレビュー & 対象選択)**:
   * 「検出されたテキストブロック一覧」コンテナ内のチェックボックスから、**スライド上で再編集可能にしたいテキスト**にチェックを入れます。
   * チェックを入れると、その部分が画像上で透過（チェッカーボード）に切り抜かれます。
2. **右カラム (仕上がりプレビュー & テキスト編集)**:
   * 切り抜かれた箇所が OpenCV によって綺麗に背景補完され、その上に新しい編集テキストが配置された「仕上がりプレビュー (Live Preview)」がリアルタイム表示されます。
   * 編集対象のドロップダウンからブロックを選択し、テキスト内容、フォントスタイル、サイズ微調整、テキストカラーに加え、**「位置の微調整（X軸・Y軸シフト）」および「配置枠の幅の拡縮」**をピクセル単位でカスタマイズできます。
3. **ダウンロード**:
   * **ローカル直接保存 (推奨)**: 「💾 ローカルPCの「ダウンロード」フォルダに直接保存する」ボタンを押すと、お使いの PC の `ダウンロード` フォルダにダイレクトに保存されます。
   * **ブラウザダウンロード**: 下部のダウンロードボタンからブラウザ経由で `.pptx`、`.png`、`.jpg` を取得することも可能です。

---

## 5. GitHub リポジトリ (GitHub Repository)

コードの管理、コントリビューション、最新バージョンの取得は以下のリポジトリで行っています：

🔗 [https://github.com/MT-Git01/Nanobanana_Editor.git](https://github.com/MT-Git01/Nanobanana_Editor.git)
