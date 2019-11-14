# Global Microsoft 365 Developer Bootcamp 2019:Office Add-ins Hands-on
[Global Microsoft 365 Developer Bootcamp 2019 Tokyo](https://connpass.com/event/144707/)、「__Office アドイン__」のハンズオン用資料です。クローン、もしくは[Download ZIP](https://glodia.jp/2017/11/13/2644/)してお使いください。

## ハンズオンの目的
本ハンズオンは、「__Office アドイン__」の概要と開発方法の学習を目的としています。
実際にアドインを動かして仕組みを理解し、`Script Lab`で遊んで __“Office アドインでどんなことができるのか”__ を体験してみましょう！

## ハンズオン環境

|  |  |
|------|-------------|
| OS | Windows 10 Pro x64 |
| Office | [Office Online](https://www.office.com/) (Office on the web), Office 365 ProPlus |
| Server | [XAMPP Portable](https://sourceforge.net/projects/xampp/files/XAMPP%20Windows/) |
| Editor | メモ帳, [Visual Studio Code](https://code.visualstudio.com/)

## ハンズオン内容

<ol>
<li>Office アドインの概要説明・アドイン紹介 (ハンズオン資料・第1章参照)</li>
<li>メモ帳とXAMPPを使ったOffice アドイン開発 (ハンズオン資料・第2章参照)</li>

> Excelアドインのテストができたら、マニフェストファイルの`Host`要素を変更して、WordやPowerPointのアドインも試してみましょう。コードは「[最初の Word アドインをビルドする](https://docs.microsoft.com/ja-jp/office/dev/add-ins/quickstarts/word-quickstart?tabs=visual-studio-code)」、「[最初の PowerPoint アドインをビルドする](https://docs.microsoft.com/ja-jp/office/dev/add-ins/quickstarts/powerpoint-quickstart?tabs=visual-studio-code)」参照

> 動作確認は[Office Online](https://www.office.com/)で行います。第7章を参考に、マニフェストファイルをアップロードしてみましょう。

<li>Node.js環境構築 (ハンズオン資料・第4章参照)</li>

> [Node.js](https://nodejs.org/ja/)をインストールして、npmからYO OFFICE!(Yeoman)をインストールしてみましょう。インストールコマンドは`npm install -g yo generator-office`です。

<li>Visual Studio Codeを使ったOffice アドイン開発 (ハンズオン資料・第8章参照)</li>

> [Visual Studio Code](https://code.visualstudio.com/)とYO OFFICE!を使ったOffice アドイン開発を体験してみましょう。

> 使用するコマンド：`yo office` , `npm run start:web`

<li>Script Lab (ハンズオン資料・第5章参照)</li>

> `Script Lab`はサンプルコードの動作確認や、自分で書いたコードのテストも簡単に行えます。Excelだけではなく、WordやPowerPointでも遊んで、Office アドインで __“できること”__ を色々と体験してみましょう！

<li>参考書籍紹介 (ハンズオン資料・第11章参照)</li>
</ol>

## もっとハンズオン！

余裕がある方は、是非下記内容にもチャレンジしてみてください！ :smile:

<ol>
<li>Excel カスタム関数(Excel Custom Functions)開発</li>

> 「[Excel でカスタム関数を作成する](https://docs.microsoft.com/ja-jp/office/dev/add-ins/excel/custom-functions-overview
)」参照

<li>Office UI Fabricによる画面設計</li>

> 「[Office アドインでの Office UI Fabric](https://docs.microsoft.com/ja-jp/office/dev/add-ins/design/office-ui-fabric)」参照

<li>SharePointカタログを使ったOffice アドインの発行</li>

> ハンズオン資料・第7章参照、および「[作業ウィンドウ アドインとコンテンツ アドインを SharePoint カタログに発行する](https://docs.microsoft.com/ja-jp/office/dev/add-ins/publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog)」参照

</ol>