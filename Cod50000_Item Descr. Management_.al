codeunit 50000 "Item Descr. Management"
{
  EventSubscriberInstance = StaticAutomatic;

  [EventSubscriber(ObjectType::Table, 370, 'OnBeforeParseCellValue', '', true, true)]
  local procedure ParseCellValue2Blob(var ExcelBuffer: Record "Excel Buffer";
  var Value: Text;
  var FormatString: Text;
  var isHandled: Boolean)var TempBlob: Record TempBlob temporary;
  begin
    CLEAR(ExcelBuffer."Cell Value as Blob");
    if StrLen(Value) > 250 then begin
      TempBlob.Blob:=ExcelBuffer."Cell Value as Blob";
      TempBlob.WriteAsText(Value, TEXTENCODING::UTF8);
      ExcelBuffer."Cell Value as Blob":=TempBlob.Blob;
      isHandled:=true;
    end;
  end;
  procedure ImportExcelSheet()var ItemDescr: Record "Item Description";
  FileName: Text;
  FileManagement: Codeunit "File Management";
  ExcelStream: InStream;
  Rows: Integer;
  Columns: Integer;
  RowNo: Integer;
  ColumnNo: Integer;
  begin
    ExcelBuffer.LockTable();
    ExcelBuffer.Reset();
    ExcelBuffer.DeleteAll();
    FileName:='';
    UploadIntoStream(ImportTitleTxt, '', FileManagement.GetToFilterText('', ExcelFileNameTxt), FileName, ExcelStream);
    if FileName = '' then exit;
    ExcelBuffer.OpenBookStream(ExcelStream, ItemDescr.TableCaption);
    ExcelBuffer.ReadSheet;
    if ExcelBuffer.FindLast then;
    Rows:=ExcelBuffer."Row No.";
    Columns:=ExcelBuffer."Column No.";
    if(Rows < 4) or (Columns < 29)then Error(WrongFormatExcelFileErr);
    for RowNo:=4 to Rows do begin
      // for ColumnNo := 1 to Columns do
      // ExcelBuffer.Get(RowNo, 1);
      // if ItemDescr.Get(ExcelBuffer."Cell Value as Text") then begin
      //     ItemDescr.SetDescription(GetValueAtCell(RowNo, 2));
      //     ItemDescr.SetBulletPoint1(GetValueAtCell(RowNo, 3));
      //     ItemDescr.SetBulletPoint2(GetValueAtCell(RowNo, 4));
      //     ItemDescr.SetBulletPoint3(GetValueAtCell(RowNo, 5));
      //     ItemDescr.SetBulletPoint4(GetValueAtCell(RowNo, 6));
      //     ItemDescr."Bullet Point 5" := GetValueAtCell(RowNo, 7);
      //     ItemDescr."Main Image URL" := GetValueAtCell(RowNo, 8);
      //     ItemDescr."Other Image URL" := GetValueAtCell(RowNo, 9);
      //     ItemDescr."Label Image URL" := GetValueAtCell(RowNo, 10);
      //     ItemDescr."Label Image URL 2" := GetValueAtCell(RowNo, 11);
      //     ItemDescr.SetSearchTerms(GetValueAtCell(RowNo, 12));
      //     ItemDescr.SetSearchTermsForGoogleOnly(GetValueAtCell(RowNo, 13));
      //     ItemDescr.SetIngredients(GetValueAtCell(RowNo, 14));
      //     ItemDescr.SetIndications(GetValueAtCell(RowNo, 15));
      //     ItemDescr.SetDirections(GetValueAtCell(RowNo, 16));
      //     ItemDescr.SetWarning(GetValueAtCell(RowNo, 17));
      //     ItemDescr."Name RU" := GetValueAtCell(RowNo, 18);
      //     ItemDescr."Name RU 2" := GetValueAtCell(RowNo, 19);
      //     ItemDescr.SetDescriptionRU(GetValueAtCell(RowNo, 20));
      //     ItemDescr.SetBulletPoint1RU(GetValueAtCell(RowNo, 21));
      //     ItemDescr.SetBulletPoint2RU(GetValueAtCell(RowNo, 22));
      //     ItemDescr.SetBulletPoint3RU(GetValueAtCell(RowNo, 23));
      //     ItemDescr.SetBulletPoint4RU(GetValueAtCell(RowNo, 24));
      //     ItemDescr."Bullet Point 5 RU" := GetValueAtCell(RowNo, 25);
      //     ItemDescr.SetIngredientsRU(GetValueAtCell(RowNo, 26));
      //     ItemDescr.SetIndicationsRU(GetValueAtCell(RowNo, 27));
      //     ItemDescr.SetDirectionsRU(GetValueAtCell(RowNo, 28));
      //     Evaluate(ItemDescr.New, GetValueAtCell(RowNo, 29));
      //     Evaluate(ItemDescr."Sell-out", GetValueAtCell(RowNo, 30));
      //     ItemDescr.Barcode := GetValueAtCell(RowNo, 31);
      //     ItemDescr.Modify();
      // end else begin
      // ItemDescr.Init();
      ItemDescr."Item No.":=GetValueAtCell(RowNo, 1);
      ItemDescr.SetDescription(GetValueAtCell(RowNo, 2));
      ItemDescr.SetBulletPoint1(GetValueAtCell(RowNo, 3));
      ItemDescr.SetBulletPoint2(GetValueAtCell(RowNo, 4));
      ItemDescr.SetBulletPoint3(GetValueAtCell(RowNo, 5));
      ItemDescr.SetBulletPoint4(GetValueAtCell(RowNo, 6));
      ItemDescr."Bullet Point 5":=GetValueAtCell(RowNo, 7);
      ItemDescr."Main Image URL":=GetValueAtCell(RowNo, 8);
      ItemDescr."Other Image URL":=GetValueAtCell(RowNo, 9);
      ItemDescr."Label Image URL":=GetValueAtCell(RowNo, 10);
      ItemDescr."Label Image URL 2":=GetValueAtCell(RowNo, 11);
      ItemDescr.SetSearchTerms(GetValueAtCell(RowNo, 12));
      ItemDescr.SetSearchTermsForGoogleOnly(GetValueAtCell(RowNo, 13));
      ItemDescr.SetIngredients(GetValueAtCell(RowNo, 14));
      ItemDescr.SetIndications(GetValueAtCell(RowNo, 15));
      ItemDescr.SetDirections(GetValueAtCell(RowNo, 16));
      ItemDescr.SetWarning(GetValueAtCell(RowNo, 17));
      ItemDescr."Name RU":=GetValueAtCell(RowNo, 18);
      ItemDescr."Name RU 2":=GetValueAtCell(RowNo, 19);
      ItemDescr.SetDescriptionRU(GetValueAtCell(RowNo, 20));
      ItemDescr.SetBulletPoint1RU(GetValueAtCell(RowNo, 21));
      ItemDescr.SetBulletPoint2RU(GetValueAtCell(RowNo, 22));
      ItemDescr.SetBulletPoint3RU(GetValueAtCell(RowNo, 23));
      ItemDescr.SetBulletPoint4RU(GetValueAtCell(RowNo, 24));
      ItemDescr."Bullet Point 5 RU":=GetValueAtCell(RowNo, 25);
      ItemDescr.SetIngredientsRU(GetValueAtCell(RowNo, 26));
      ItemDescr.SetIndicationsRU(GetValueAtCell(RowNo, 27));
      ItemDescr.SetDirectionsRU(GetValueAtCell(RowNo, 28));
      Evaluate(ItemDescr.New, GetValueAtCell(RowNo, 29));
      Evaluate(ItemDescr."Sell-out", GetValueAtCell(RowNo, 30));
      ItemDescr.Barcode:=GetValueAtCell(RowNo, 31);
      if ItemDescr.Insert()then ItemDescr.Modify();
    // end;
    end;
    ExcelBuffer.CloseBook;
    ExcelBuffer.DeleteAll();
  end;
  procedure GetValueAtCell(RowNo: Integer;
  ColNo: Integer): Text;
  var TempBlob: Record TempBlob;
  CR: Text[1];
  begin
    if not ExcelBuffer.Get(RowNo, ColNo)then exit('');
    if not ExcelBuffer."Cell Value as Blob".HasValue then exit(ExcelBuffer."Cell Value as Text");
    ExcelBuffer.CALCFIELDS("Cell Value as Blob");
    CR[1]:=10;
    TempBlob.Blob:=ExcelBuffer."Cell Value as Blob";
    EXIT(TempBlob.ReadAsText(CR, TEXTENCODING::UTF8));
  end;
  var ExcelBuffer: Record "Excel Buffer";
  ImportTitleTxt: Label 'Choose the Excel worksheet where data classifications have been added.';
  ExcelFileNameTxt: Label '*.xlsx';
  WrongFormatExcelFileErr: Label 'Looks like the Excel worksheet you provided is not formatted correctly.';
}
