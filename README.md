# Global Office 365 Developer Bootcamp:Office Add-ins Hands-on
[Global Office 365 Developer Bootcamp 2018(Japan)](https://connpass.com/event/91901/)、「__Office アドイン__」のハンズオン用資料です。

## ハンズオン環境

|  |  |
|------|-------------|
| OS | Windows 10 Pro x64 |
| Office | Office 365 ProPlus バージョン1803(主にExcel使用) |
| Server | [XAMPP Portable](https://sourceforge.net/projects/xampp/files/XAMPP%20Windows/) |
| Editor | メモ帳など

## ハンズオン手順

<ol>
<li>Office アドインの概要 (ハンズオン資料・第1章参照)</li>
<li>メモ帳とXAMPPを使ったOffice アドイン開発 (ハンズオン資料・第2章参照)</li>

> Excelアドインのテストができたら、マニフェストファイルの`Host`要素を変更して、WordやPowerPointのアドインも試してみましょう。コードは「[最初の Word アドインをビルドする](https://docs.microsoft.com/ja-jp/office/dev/add-ins/quickstarts/word-quickstart?tabs=visual-studio-code)」、「[最初の PowerPoint アドインをビルドする](https://docs.microsoft.com/ja-jp/office/dev/add-ins/quickstarts/powerpoint-quickstart?tabs=visual-studio-code)」参照

<li>アドイン コマンド (ハンズオン資料・第3章参照)</li>
<li>YO OFFICE!(Yeoman)によるひな形の作成 (ハンズオン資料・第4章参照)</li>
<li>Script Lab (ハンズオン資料・第5章参照)</li>

> `Script Lab`はサンプルコードの動作確認や、自分で書いたコードのテストも簡単に行えます。Excelだけではなく、WordやPowerPointでも遊んで、Office アドインで __“できること”__ を色々と体験してみましょう！

<li>(Insider環境であれば)Excel カスタム関数(Excel Custom Functions) (ハンズオン資料・第6章参照)</li>
</ol>

## もっとハンズオン！

余裕がある方は、是非下記内容にもチャレンジしてみてください！ :smile:

<ol>
<li>「Visual Studio Code」によるOffice アドイン開発。</li>

> 「[Office Add-ins with Visual Studio Code](https://code.visualstudio.com/docs/other/office)」参照、TypeScriptの型定義ファイル：[https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js)

<li>Office UI Fabricによる画面設計</li>

> 「[Office アドインでの Office UI Fabric](https://docs.microsoft.com/ja-jp/office/dev/add-ins/design/office-ui-fabric)」参照

<li>SharePointカタログを使ったOffice アドインの発行</li>

> 「[作業ウィンドウ アドインとコンテンツ アドインを SharePoint カタログに発行する](https://docs.microsoft.com/ja-jp/office/dev/add-ins/publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog)」参照

</ol>