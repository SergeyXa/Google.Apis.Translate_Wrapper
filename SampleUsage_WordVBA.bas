' This sample shows how to use the translate wrapper in Microsoft Word. The macro
' translated the entire document from French to English. The translated text is inserted
' right after the original text, in blue.

Attribute VB_Name = "Module1"
Sub InternationalizeDoc()

    Set Translator = CreateObject("GoogleApisTranslateWrapper.Translator")
    
    Set doc = ActiveDocument
    doc.Content.InsertParagraphAfter
    
    Dim para As Range, newPara As Range
    
    Set para = doc.Paragraphs(1).Range
    Do
        If (para.End - para.Start > 2) And (Trim(Replace(Replace(para.Text, vbCr, ""), vbLf, "")) <> "") Then
            translatedText = Translator.Translate(para.Text, "fr", "en")
        
            Set newPara = doc.Range(para.End - 1, para.End - 1)
            newPara.InsertAfter Chr(11)
            
            If newPara.Paragraphs.Alignment = wdAlignParagraphJustify Then _
                newPara.Paragraphs.Alignment = wdAlignParagraphLeft
            
            Set newPara = doc.Range(newPara.End, newPara.End)
            newPara.Text = translatedText
            newPara.Font.Size = newPara.Font.Size - 2
            newPara.Font.ColorIndex = wdDarkBlue
            
            Set para = newPara.Next(wdParagraph)
        Else
            Set para = para.Next(wdParagraph)
        End If
    Loop Until para Is Nothing
End Sub


