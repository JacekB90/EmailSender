
Public Class EmailSender

    Public Shared Sub Main()

    End Sub

    Public Sub NotesMailSend(strOdbiorca As String, strTemat As String,
    strTresc As String, strFilename As String)

        Dim objNotes As Object, objNotesDB As Object, objNotesMailDoc As Object
        Dim SendItem, NCopyItem, BlindCopyToItem, rtitem


        objNotes = GetObject("", "Notes.Notessession")
        objNotesDB = objNotes.GETDATABASE("", "")

        Call objNotesDB.OPENMAIL
        objNotesMailDoc = objNotesDB.CREATEDOCUMENT
        objNotesMailDoc.Form = "Memo"
        Call objNotesMailDoc.Save(True, False)
        SendItem = objNotesMailDoc.APPENDITEMVALUE("SendTo", "")
        NCopyItem = objNotesMailDoc.APPENDITEMVALUE("CopyTo", "")
        BlindCopyToItem = objNotesMailDoc.APPENDITEMVALUE("BlindCopyTo", "")
        objNotesMailDoc.sendto = strOdbiorca
        objNotesMailDoc.Subject = strTemat
        rtitem = objNotesMailDoc.CREATERICHTEXTITEM("Body")
        objNotesMailDoc.Body = strTresc
        rtitem.ADDNEWLINE(1)
        Call rtitem.EMBEDOBJECT(1454, "", strFilename)

        Call objNotesMailDoc.Save(True, False)
        Call objNotesMailDoc.Send(False)
        objNotesMailDoc.RemoveItem("DeliveredDate")
        Call objNotesMailDoc.Save(True, False)

        Call objNotes.Close

        objNotes = Nothing

    End Sub

End Class

