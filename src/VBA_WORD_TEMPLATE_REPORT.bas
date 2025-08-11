Sub InsertReportStructure()
    ' --- カーソル位置にレポート構成を挿入 ---
    Dim sel As Selection
    Set sel = Application.Selection

    ' 各見出しを挿入
    Call InsertHeadingAndParagraph(sel, "1. 実験目的")
    Call InsertHeadingAndParagraph(sel, "2. 実験原理")
    Call InsertHeadingAndParagraph(sel, "3. 実験結果")
    Call InsertHeadingAndParagraph(sel, "4. 考察")
    Call InsertHeadingAndParagraph(sel, "5. 結論")
    Call InsertHeadingAndParagraph(sel, "6. 検討事項")

    ' 参考文献は空行を挿入しない
    sel.Style = ActiveDocument.Styles("見出し 1")
    sel.TypeText Text:="7. 参考文献"
    sel.TypeParagraph()
    sel.Style = ActiveDocument.Styles("標準")
    ' 参考文献は段落番号を使うため、手動で設定するか、
    ' ListGalleries(wdNumberGallery).ListTemplates(1).Name = "" のようなコードで設定します。

    MsgBox "レポートの構成を挿入しました。"
End Sub

' --- ヘルパーサブルーチン ---
Private Sub InsertHeadingAndParagraph(ByVal sel As Selection, ByVal headingText As String)
    ' 見出しを挿入
    sel.Style = ActiveDocument.Styles("見出し 1")
    sel.TypeText Text:=headingText
    sel.TypeParagraph() ' 改行

    ' 本文用の空行を挿入
    sel.Style = ActiveDocument.Styles("標準")
    sel.TypeParagraph()
End Sub
