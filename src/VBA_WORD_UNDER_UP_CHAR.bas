' 一時文字の定義（ドキュメント内で使用されないであろう私用領域の文字を利用）
Private Const TEMP_UNDERSCORE As String = ChrW(&HE000)
Private Const TEMP_CARET As String = ChrW(&HE001)

' =================================================================
' メインマクロ：これを実行すればOK
' =================================================================
Sub ConvertMarkupText()
    ' 実行中の画面更新を停止して処理を高速化
    Application.ScreenUpdating = False

    ' 1. エスケープ文字を一時的な文字に置換
    Call EscapeMarkers
    
    ' 2. 下付き文字、上付き文字に変換
    Call ConvertToSubscript
    Call ConvertToSuperscript
    
    ' 3. 一時的な文字を元に戻す
    Call RestoreEscapedMarkers

    ' 画面更新を再開
    Application.ScreenUpdating = True
    
    MsgBox "マークアップの変換が完了しました。"
End Sub


' =================================================================
' ヘルパーマクロ：個別の処理
' =================================================================

' --- 下付き文字の変換 (`_text_` -> text) ---
Private Sub ConvertToSubscript()
    Dim rng As Range
    Set rng = ActiveDocument.Content
    
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        ' ワイルドカード検索: _ で囲まれた1文字以上の文字列にマッチ
        ' <*>,@, ?などのワイルドカードが使用可能
        .Text = "_(<*>)_"
        .MatchWildcards = True
        .Wrap = wdFindStop
        
        Do While .Execute
            ' 前後の "_" を削除
            rng.Characters.Last.Delete
            rng.Characters.First.Delete
            
            ' 残った範囲の文字を下付きに設定
            rng.Font.Subscript = True

            ' 下付き/上付き設定時に他の書式（太字など）が解除されることがあるため、
            ' 現在の書式を再適用してこれを防ぐ
            rng.Font.Bold = rng.Font.Bold
            rng.Font.Italic = rng.Font.Italic

            ' 検索範囲を処理済みの箇所の直後に移動してループを継続
            rng.Collapse wdCollapseEnd
        Loop
    End With
End Sub

' --- 上付き文字の変換 (`^text^` -> text) ---
Private Sub ConvertToSuperscript()
    Dim rng As Range
    Set rng = ActiveDocument.Content
    
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        ' ワイルドカード検索: ^ で囲まれた1文字以上の文字列にマッチ
        ' \^ とすることで、^をワイルドカードではなく通常の文字として検索
        .Text = "\^(<*>)\^"
        .MatchWildcards = True
        .Wrap = wdFindStop
        
        Do While .Execute
            ' 前後の "^" を削除
            rng.Characters.Last.Delete
            rng.Characters.First.Delete
            
            ' 残った範囲の文字を上付きに設定
            rng.Font.Superscript = True
            ' 書式維持のため、現在の書式を再適用
            rng.Font.Bold = rng.Font.Bold
            rng.Font.Italic = rng.Font.Italic

            ' 検索範囲を処理済みの箇所の直後に移動してループを継続
            rng.Collapse wdCollapseEnd
        Loop
    End With
End Sub

' --- エスケープ処理 (準備) ---
Private Sub EscapeMarkers()
    ' Wordの通常の置換機能を使用
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = False
        
        ' "__" を一時文字に置換
        .Text = "__"
        .Replacement.Text = TEMP_UNDERSCORE
        .Execute Replace:=wdReplaceAll
        
        ' "^^" を一時文字に置換
        .Text = "^^"
        .Replacement.Text = TEMP_CARET
        .Execute Replace:=wdReplaceAll
    End With
End Sub

' --- エスケープ処理 (復元) ---
Private Sub RestoreEscapedMarkers()
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = False
        
        ' 一時文字を "_" に戻す
        .Text = TEMP_UNDERSCORE
        .Replacement.Text = "_"
        .Execute Replace:=wdReplaceAll
        
        ' 一時文字を "^" に戻す
        .Text = TEMP_CARET
        .Replacement.Text = "^"
        .Execute Replace:=wdReplaceAll
    End With
End Sub
