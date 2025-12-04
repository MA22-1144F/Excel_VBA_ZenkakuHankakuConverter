# Excel_VBA_ZenkakuHankakuConverter

## 概要
Excelの選択範囲のテキストを全角から半角または半角から半角に変換します。  
PERSONAL.XLSBに保存して動作させることを想定しています。  

<img src="image_zenhancon.png" alt="イメージ画面" width="600">

## 動作環境
Microsoft Excel上で動作します。  

## インストール方法
1. Contentsフォルダ内の frmZenkakuHankaku.frm、frmZenkakuHankaku.frx、全角半角変換.bas を任意の同一フォルダに保存
2. Excelで新規WorkSheetを開く
3. 開発タブのVisual BasicまたはAlt+F11でVBE(Visual Basic Editor)を開く
4. VBAProject一覧から VBAProject (PERSONAL.XLSB) を選択し右クリック
5. ファイルのインポートで保存した frmZenkakuHankaku.frm を開く
6. 加えてファイルのインポートで保存した 全角半角変換.bas を開く
7. 上書き保存ボタンまたはCtrl+Sで保存
以下オプション設定  
8. 任意のExcel WorkSheetに戻り、ファイル→オプション→リボンのユーザー設定の順で遷移
9. 任意のユーザー設定グループ(無ければ新規作成)に PERSONAL.XLSB!ZenkakuHankakuConverter のマクロを追加
10. お好みで名前やアイコンを変更
11. Excel WorkSheetに戻り、設定したマクロのアイコンを選択して起動

## 機能
選択範囲について以下の機能を実行

* 変換方向の選択  
  全角→半角 または 半角→全角  

* 変換対象の選択  
  * 英数字  
  * 記号    
  * カタカナ  
  * スペース  

* 数式が入力されたセルに対する処理の選択

## ライセンス
MIT License
