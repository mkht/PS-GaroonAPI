﻿#
# モジュール 'PS-GaroonAPI' のモジュール マニフェスト
#
# 生成者: mkht
#
# 生成日: 2016/08/23
#

@{

    # このマニフェストに関連付けられているスクリプト モジュール ファイルまたはバイナリ モジュール ファイル。
    RootModule        = 'PS-GaroonAPI.psm1'

    # このモジュールのバージョン番号です。
    ModuleVersion     = '0.4.0'

    # サポートされている PSEditions
    # CompatiblePSEditions = @()

    # このモジュールを一意に識別するために使用される ID
    GUID              = 'e23a58c9-cb72-411a-91eb-464b5f4ab3d7'

    # このモジュールの作成者
    Author            = 'mkht'

    # このモジュールの会社またはベンダー
    CompanyName       = 'mkht'

    # このモジュールの著作権情報
    Copyright         = '(c) 2021 mkht. All rights reserved.'

    # このモジュールの機能の説明
    Description       = 'PowerShell Garoon API Module'

    # このモジュールに必要な Windows PowerShell エンジンの最小バージョン
    PowerShellVersion = '5.0'

    # このモジュールに必要な Windows PowerShell ホストの名前
    # PowerShellHostName = ''

    # このモジュールに必要な Windows PowerShell ホストの最小バージョン
    # PowerShellHostVersion = ''

    # Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
    # DotNetFrameworkVersion = ''

    # Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
    # CLRVersion = ''

    # このモジュールに必要なプロセッサ アーキテクチャ (なし、X86、Amd64)
    # ProcessorArchitecture = ''

    # このモジュールをインポートする前にグローバル環境にインポートされている必要があるモジュール
    # RequiredModules = @()

    # このモジュールをインポートする前に読み込まれている必要があるアセンブリ
    # RequiredAssemblies = @()

    # このモジュールをインポートする前に呼び出し元の環境で実行されるスクリプト ファイル (.ps1)。
    ScriptsToProcess  = @(
        'Class\GaroonClass.ps1',
        'Class\GaroonBase.ps1',
        'Class\GaroonAdmin.ps1',
        'Class\GaroonMail.ps1',
        'Class\GaroonAddress.ps1'
    )

    # このモジュールをインポートするときに読み込まれる型ファイル (.ps1xml)
    # TypesToProcess = @()

    # このモジュールをインポートするときに読み込まれる書式ファイル (.ps1xml)
    # FormatsToProcess = @()

    # RootModule/ModuleToProcess に指定されているモジュールの入れ子になったモジュールとしてインポートするモジュール
    # NestedModules = @()

    # このモジュールからエクスポートする関数です。最適なパフォーマンスを得るには、ワイルドカードを使用せず、エクスポートする関数がない場合は、エントリを削除しないで空の配列を使用してください。
    FunctionsToExport = @(
        'Get-GrnUser',
        'New-GrnUser',
        'Set-GrnUser',
        'Remove-GrnUser',
        'Get-GrnOrganization',
        'New-GrnOrganization',
        'Set-GrnOrganization',
        'Add-GrnOrganizationMember',
        'Remove-GrnOrganizationMember',
        'New-GrnMailAccount',
        'Get-GrnAddressBook',
        'Add-GrnAddressBookMember',
        'Set-GrnAddressBookMember',
        'Remove-GrnAddressBookMember',
        'Invoke-SOAPRequest'
    )

    # このモジュールからエクスポートするコマンドレットです。最適なパフォーマンスを得るには、ワイルドカードを使用せず、エクスポートするコマンドレットがない場合は、エントリを削除しないで空の配列を使用してください。
    CmdletsToExport   = @()

    # このモジュールからエクスポートする変数
    # VariablesToExport = '*'

    # このモジュールからエクスポートするエイリアスです。最適なパフォーマンスを得るには、ワイルドカードを使用せず、エクスポートするエイリアスがない場合は、エントリを削除しないで空の配列を使用してください。
    AliasesToExport   = @()

    # このモジュールからエクスポートする DSC リソース
    # DscResourcesToExport = @()

    # このモジュールに同梱されているすべてのモジュールのリスト
    # ModuleList = @()

    # このモジュールに同梱されているすべてのファイルのリスト
    # FileList = @()

    # RootModule/ModuleToProcess に指定されているモジュールに渡すプライベート データ。これには、PowerShell で使用される追加のモジュール メタデータを含む PSData ハッシュテーブルが含まれる場合もあります。
    PrivateData       = @{

        PSData = @{

            # このモジュールに適用されているタグ。オンライン ギャラリーでモジュールを検出する際に役立ちます。
            # Tags = @()

            # このモジュールのライセンスの URL。
            # LicenseUri = ''

            # このプロジェクトのメイン Web サイトの URL。
            # ProjectUri = ''

            # このモジュールを表すアイコンの URL。
            # IconUri = ''

            # このモジュールの ReleaseNotes
            # ReleaseNotes = ''

        } # PSData ハッシュテーブル終了

    } # PrivateData ハッシュテーブル終了

    # このモジュールの HelpInfo URI
    # HelpInfoURI = ''

    # このモジュールからエクスポートされたコマンドの既定のプレフィックス。既定のプレフィックスをオーバーライドする場合は、Import-Module -Prefix を使用します。
    # DefaultCommandPrefix = ''

}

