### README

#### Mod_Format_Range (Module)

**Description**

This module provides utilities for handling range formatting in Excel VBA. It includes functions and subroutines to work with cell ranges, format them based on specific criteria, and determine array properties.

**Index**
- **Public Type Min_And_Max**: User-defined type for storing minimum and maximum values (e.g., row/column indices).
- **Public Function Get_Max**: Retrieves the maximum row or column index from the specified base cell.
- **Public Function Is_Array_Empty**: Returns `True` if the array has no elements.
- **Public Sub Set_Format**: Applies formatting settings to a specified range based on keywords found in the header row.
- **Private Sub Set_Format_Code**: Auxiliary processing to apply specific formatting to a range.

**History**
- **2024/09/11 (Ver.1.0.0)**: Created as new module. [Kyosuke Homma, https://github.com/kyosukehomma]

**Developer**
- **Name**: Kyosuke Homma
- **Contact**: kyosukehomma@gmail.com

---

### Japanese description follows below.

#### Mod_Format_Range (モジュール)

**説明**

このモジュールは、Excel VBAでのセル範囲のフォーマット処理に関するユーティリティを提供します。特定の基準に基づいてセル範囲をフォーマットしたり、配列のプロパティを判断するための関数とサブルーチンが含まれています。

**インデックス**
- **Public Type Min_And_Max**: 最小値と最大値を格納するユーザー定義型（例: 行/列インデックス）。
- **Public Function Get_Max**: 指定された基準セルから最大の行または列インデックスを取得します。
- **Public Function Is_Array_Empty**: 配列に要素が存在しない場合に `True` を返します。
- **Public Sub Set_Format**: ヘッダー行に見つかったキーワードに基づいて、指定された範囲にフォーマット設定を適用します。
- **Private Sub Set_Format_Code**: 特定のフォーマットを範囲に適用するための補助処理です。

**履歴**
- **2024/09/11 (Ver.1.0.0)**: 新しいモジュールとして作成されました。 [Kyosuke Homma, https://github.com/kyosukehomma]

**開発者**
- **名前**: Kyosuke Homma
- **連絡先**: kyosukehomma@gmail.com
