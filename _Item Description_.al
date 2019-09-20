table 50000 "Item Description"
{
  DataClassification = ToBeClassified;
  Caption = 'Item Description';

  fields
  {
    field(1;"Item No.";Code[20])
    {
      DataClassification = ToBeClassified;
      TableRelation = Item."No.";
    }
    field(2;Description;Blob)
    {
      DataClassification = ToBeClassified;
      Subtype = Memo;
    }
    field(3;"Bullet Point 1";Blob)
    {
      DataClassification = ToBeClassified;
      Subtype = Memo;
    }
    field(4;"Bullet Point 2";Blob)
    {
      DataClassification = ToBeClassified;
      Subtype = Memo;
    }
    field(5;"Bullet Point 3";Blob)
    {
      DataClassification = ToBeClassified;
      Subtype = Memo;
    }
    field(6;"Bullet Point 4";Blob)
    {
      DataClassification = ToBeClassified;
      Subtype = Memo;
    }
    field(7;"Bullet Point 5";Text[100])
    {
      DataClassification = ToBeClassified;
    }
    field(8;"Main Image URL";Text[150])
    {
      DataClassification = ToBeClassified;
    }
    field(9;"Other Image URL";Text[150])
    {
      DataClassification = ToBeClassified;
    }
    field(10;"Label Image URL";Text[150])
    {
      DataClassification = ToBeClassified;
    }
    field(11;"Label Image URL 2";Text[150])
    {
      DataClassification = ToBeClassified;
    }
    field(12;"Search Terms";Blob)
    {
      DataClassification = ToBeClassified;
      Subtype = Memo;
    }
    field(13;"Search Terms for Google only";Blob)
    {
      DataClassification = ToBeClassified;
      Subtype = Memo;
    }
    field(14;Ingredients;Blob)
    {
      DataClassification = ToBeClassified;
      Subtype = Memo;
    }
    field(15;Indications;Blob)
    {
      DataClassification = ToBeClassified;
      Subtype = Memo;
    }
    field(16;Directions;Blob)
    {
      DataClassification = ToBeClassified;
      Subtype = Memo;
    }
    field(17;Warning;Blob)
    {
      DataClassification = ToBeClassified;
      Subtype = Memo;
    }
    field(18;"Name RU";Text[100])
    {
      DataClassification = ToBeClassified;
    }
    field(19;"Name RU 2";Text[50])
    {
      DataClassification = ToBeClassified;
    }
    field(20;"Description RU";Blob)
    {
      DataClassification = ToBeClassified;
      Subtype = Memo;
    }
    field(21;"Bullet Point 1 RU";Blob)
    {
      DataClassification = ToBeClassified;
      Subtype = Memo;
    }
    field(22;"Bullet Point 2 RU";Blob)
    {
      DataClassification = ToBeClassified;
      Subtype = Memo;
    }
    field(23;"Bullet Point 3 RU";Blob)
    {
      DataClassification = ToBeClassified;
      Subtype = Memo;
    }
    field(24;"Bullet Point 4 RU";Blob)
    {
      DataClassification = ToBeClassified;
      Subtype = Memo;
    }
    field(25;"Bullet Point 5 RU";Text[100])
    {
      DataClassification = ToBeClassified;
    }
    field(26;"Ingredients RU";Blob)
    {
      DataClassification = ToBeClassified;
      Subtype = Memo;
    }
    field(27;"Indications RU";Blob)
    {
      DataClassification = ToBeClassified;
      Subtype = Memo;
    }
    field(28;"Directions RU";Blob)
    {
      DataClassification = ToBeClassified;
      Subtype = Memo;
    }
    field(29;"New";Date)
    {
      DataClassification = ToBeClassified;
    }
    field(30;"Sell-out";Date)
    {
      DataClassification = ToBeClassified;
    }
    field(31;Barcode;Text[20])
    {
      DataClassification = ToBeClassified;
    }
  }
  keys
  {
    key(PK;"Item No.")
    {
      Clustered = true;
    }
  }
  var myInt: Integer;
  trigger OnInsert()begin
  end;
  trigger OnModify()begin
  end;
  trigger OnDelete()begin
  end;
  trigger OnRename()begin
  end;
  procedure InitItemDescription(ItemNo: Code[20])begin
    Init();
    "Item No.":=ItemNo;
    if not Insert()then Modify();
  end;
  // Warning
  procedure SetWarning(newWarning: Text)var TempBlob: Record TempBlob temporary;
  begin
    CLEAR(Warning);
    IF newWarning = '' THEN EXIT;
    TempBlob.Blob:=Warning;
    TempBlob.WriteAsText(newWarning, TEXTENCODING::UTF8);
    Warning:=TempBlob.Blob;
    if MODIFY then;
  end;
  procedure GetWarning(): Text;
  begin
    EXIT(GetWarningCalculated);
  end;
  local procedure GetWarningCalculated(): Text;
  var TempBlob: Record TempBlob temporary;
  CR: text[1];
  begin
    IF NOT Warning.HASVALUE THEN EXIT('');
    CALCFIELDS(Warning);
    CR[1]:=10;
    TempBlob.Blob:=Warning;
    EXIT(TempBlob.ReadAsText(CR, TEXTENCODING::UTF8));
  end;
  // Description
  procedure SetDescription(newDescription: Text)var TempBlob: Record TempBlob temporary;
  begin
    CLEAR(Description);
    IF newDescription = '' THEN EXIT;
    TempBlob.Blob:=Description;
    TempBlob.WriteAsText(newDescription, TEXTENCODING::UTF8);
    Description:=TempBlob.Blob;
    if MODIFY then;
  end;
  procedure GetDescription(): Text;
  begin
    EXIT(GetDescriptionCalculated);
  end;
  local procedure GetDescriptionCalculated(): Text;
  var TempBlob: Record TempBlob temporary;
  CR: text[1];
  begin
    IF NOT Description.HASVALUE THEN EXIT('');
    CALCFIELDS(Description);
    CR[1]:=10;
    TempBlob.Blob:=Description;
    EXIT(TempBlob.ReadAsText(CR, TEXTENCODING::UTF8));
  end;
  // BulletPoint1
  procedure SetBulletPoint1(newBulletPoint1: Text)var TempBlob: Record TempBlob temporary;
  begin
    CLEAR("Bullet Point 1");
    IF newBulletPoint1 = '' THEN EXIT;
    TempBlob.Blob:="Bullet Point 1";
    TempBlob.WriteAsText(newBulletPoint1, TEXTENCODING::UTF8);
    "Bullet Point 1":=TempBlob.Blob;
    if MODIFY then;
  end;
  procedure GetBulletPoint1(): Text;
  begin
    EXIT(GetBulletPoint1Calculated);
  end;
  local procedure GetBulletPoint1Calculated(): Text;
  var TempBlob: Record TempBlob temporary;
  CR: text[1];
  begin
    IF NOT "Bullet Point 1".HASVALUE THEN EXIT('');
    CALCFIELDS("Bullet Point 1");
    CR[1]:=10;
    TempBlob.Blob:="Bullet Point 1";
    EXIT(TempBlob.ReadAsText(CR, TEXTENCODING::UTF8));
  end;
  // BulletPoint2;
  procedure SetBulletPoint2(newBulletPoint2: Text)var TempBlob: Record TempBlob temporary;
  begin
    CLEAR("Bullet Point 2");
    IF newBulletPoint2 = '' THEN EXIT;
    TempBlob.Blob:="Bullet Point 2";
    TempBlob.WriteAsText(newBulletPoint2, TEXTENCODING::UTF8);
    "Bullet Point 2":=TempBlob.Blob;
    if MODIFY then;
  end;
  procedure GetBulletPoint2(): Text;
  begin
    EXIT(GetBulletPoint2Calculated);
  end;
  local procedure GetBulletPoint2Calculated(): Text;
  var TempBlob: Record TempBlob temporary;
  CR: text[1];
  begin
    IF NOT "Bullet Point 2".HASVALUE THEN EXIT('');
    CALCFIELDS("Bullet Point 2");
    CR[1]:=10;
    TempBlob.Blob:="Bullet Point 2";
    EXIT(TempBlob.ReadAsText(CR, TEXTENCODING::UTF8));
  end;
  // BulletPoint3
  procedure SetBulletPoint3(newBulletPoint3: Text)var TempBlob: Record TempBlob temporary;
  begin
    CLEAR("Bullet Point 3");
    IF newBulletPoint3 = '' THEN EXIT;
    TempBlob.Blob:="Bullet Point 3";
    TempBlob.WriteAsText(newBulletPoint3, TEXTENCODING::UTF8);
    "Bullet Point 3":=TempBlob.Blob;
    if MODIFY then;
  end;
  procedure GetBulletPoint3(): Text;
  begin
    EXIT(GetBulletPoint3Calculated);
  end;
  local procedure GetBulletPoint3Calculated(): Text;
  var TempBlob: Record TempBlob temporary;
  CR: text[1];
  begin
    IF NOT "Bullet Point 3".HASVALUE THEN EXIT('');
    CALCFIELDS("Bullet Point 3");
    CR[1]:=10;
    TempBlob.Blob:="Bullet Point 3";
    EXIT(TempBlob.ReadAsText(CR, TEXTENCODING::UTF8));
  end;
  // BulletPoint4
  procedure SetBulletPoint4(newBulletPoint4: Text)var TempBlob: Record TempBlob temporary;
  begin
    CLEAR("Bullet Point 4");
    IF newBulletPoint4 = '' THEN EXIT;
    TempBlob.Blob:="Bullet Point 4";
    TempBlob.WriteAsText(newBulletPoint4, TEXTENCODING::UTF8);
    "Bullet Point 4":=TempBlob.Blob;
    if MODIFY then;
  end;
  procedure GetBulletPoint4(): Text;
  begin
    EXIT(GetBulletPoint4Calculated);
  end;
  local procedure GetBulletPoint4Calculated(): Text;
  var TempBlob: Record TempBlob temporary;
  CR: text[1];
  begin
    IF NOT "Bullet Point 4".HASVALUE THEN EXIT('');
    CALCFIELDS("Bullet Point 4");
    CR[1]:=10;
    TempBlob.Blob:="Bullet Point 4";
    EXIT(TempBlob.ReadAsText(CR, TEXTENCODING::UTF8));
  end;
  // SearchTerms;
  procedure SetSearchTerms(newSearchTerms: Text)var TempBlob: Record TempBlob temporary;
  begin
    CLEAR("Search Terms");
    IF newSearchTerms = '' THEN EXIT;
    TempBlob.Blob:="Search Terms";
    TempBlob.WriteAsText(newSearchTerms, TEXTENCODING::UTF8);
    "Search Terms":=TempBlob.Blob;
    if MODIFY then;
  end;
  procedure GetSearchTerms(): Text;
  begin
    EXIT(GetSearchTermsCalculated);
  end;
  local procedure GetSearchTermsCalculated(): Text;
  var TempBlob: Record TempBlob temporary;
  CR: text[1];
  begin
    IF NOT "Search Terms".HASVALUE THEN EXIT('');
    CALCFIELDS("Search Terms");
    CR[1]:=10;
    TempBlob.Blob:="Search Terms";
    EXIT(TempBlob.ReadAsText(CR, TEXTENCODING::UTF8));
  end;
  // SearchTermsForGoogleOnly := GetSearchTermsForGoogleOnly;
  procedure SetSearchTermsForGoogleOnly(newSearchTermsForGoogleOnly: Text)var TempBlob: Record TempBlob temporary;
  begin
    CLEAR("Search Terms for Google only");
    IF newSearchTermsForGoogleOnly = '' THEN EXIT;
    TempBlob.Blob:="Search Terms for Google only";
    TempBlob.WriteAsText(newSearchTermsForGoogleOnly, TEXTENCODING::UTF8);
    "Search Terms for Google only":=TempBlob.Blob;
    if MODIFY then;
  end;
  procedure GetSearchTermsForGoogleOnly(): Text;
  begin
    EXIT(GetSearchTermsForGoogleOnlyCalculated);
  end;
  local procedure GetSearchTermsForGoogleOnlyCalculated(): Text;
  var TempBlob: Record TempBlob temporary;
  CR: text[1];
  begin
    IF NOT "Search Terms for Google only".HASVALUE THEN EXIT('');
    CALCFIELDS("Search Terms for Google only");
    CR[1]:=10;
    TempBlob.Blob:="Search Terms for Google only";
    EXIT(TempBlob.ReadAsText(CR, TEXTENCODING::UTF8));
  end;
  // Ingredients
  procedure SetIngredients(newIngredients: Text)var TempBlob: Record TempBlob temporary;
  begin
    CLEAR(Ingredients);
    IF newIngredients = '' THEN EXIT;
    TempBlob.Blob:=Ingredients;
    TempBlob.WriteAsText(newIngredients, TEXTENCODING::UTF8);
    Ingredients:=TempBlob.Blob;
    if MODIFY then;
  end;
  procedure GetIngredients(): Text;
  begin
    EXIT(GetIngredientsCalculated);
  end;
  local procedure GetIngredientsCalculated(): Text;
  var TempBlob: Record TempBlob temporary;
  CR: text[1];
  begin
    IF NOT Ingredients.HASVALUE THEN EXIT('');
    CALCFIELDS(Ingredients);
    CR[1]:=10;
    TempBlob.Blob:=Ingredients;
    EXIT(TempBlob.ReadAsText(CR, TEXTENCODING::UTF8));
  end;
  // Indications
  procedure SetIndications(newIndications: Text)var TempBlob: Record TempBlob temporary;
  begin
    CLEAR(Indications);
    IF newIndications = '' THEN EXIT;
    TempBlob.Blob:=Indications;
    TempBlob.WriteAsText(newIndications, TEXTENCODING::UTF8);
    Indications:=TempBlob.Blob;
    if MODIFY then;
  end;
  procedure GetIndications(): Text;
  begin
    EXIT(GetIndicationsCalculated);
  end;
  local procedure GetIndicationsCalculated(): Text;
  var TempBlob: Record TempBlob temporary;
  CR: text[1];
  begin
    IF NOT Indications.HASVALUE THEN EXIT('');
    CALCFIELDS(Indications);
    CR[1]:=10;
    TempBlob.Blob:=Indications;
    EXIT(TempBlob.ReadAsText(CR, TEXTENCODING::UTF8));
  end;
  // Directions
  procedure SetDirections(newDirections: Text)var TempBlob: Record TempBlob temporary;
  begin
    CLEAR(Directions);
    IF newDirections = '' THEN EXIT;
    TempBlob.Blob:=Directions;
    TempBlob.WriteAsText(newDirections, TEXTENCODING::UTF8);
    Directions:=TempBlob.Blob;
    if MODIFY then;
  end;
  procedure GetDirections(): Text;
  begin
    EXIT(GetDirectionsCalculated);
  end;
  local procedure GetDirectionsCalculated(): Text;
  var TempBlob: Record TempBlob temporary;
  CR: text[1];
  begin
    IF NOT Directions.HASVALUE THEN EXIT('');
    CALCFIELDS(Directions);
    CR[1]:=10;
    TempBlob.Blob:=Directions;
    EXIT(TempBlob.ReadAsText(CR, TEXTENCODING::UTF8));
  end;
  // DescriptionRU
  procedure SetDescriptionRU(newDescriptionRU: Text)var TempBlob: Record TempBlob temporary;
  begin
    CLEAR("Description RU");
    IF newDescriptionRU = '' THEN EXIT;
    TempBlob.Blob:="Description RU";
    TempBlob.WriteAsText(newDescriptionRU, TEXTENCODING::UTF8);
    "Description RU":=TempBlob.Blob;
    if MODIFY then;
  end;
  procedure GetDescriptionRU(): Text;
  begin
    EXIT(GetnewDescriptionRUCalculated);
  end;
  local procedure GetnewDescriptionRUCalculated(): Text;
  var TempBlob: Record TempBlob temporary;
  CR: text[1];
  begin
    IF NOT "Description RU".HASVALUE THEN EXIT('');
    CALCFIELDS("Description RU");
    CR[1]:=10;
    TempBlob.Blob:="Description RU";
    EXIT(TempBlob.ReadAsText(CR, TEXTENCODING::UTF8));
  end;
  // BulletPoint1RU
  procedure SetBulletPoint1RU(newBulletPoint1RU: Text)var TempBlob: Record TempBlob temporary;
  begin
    CLEAR("Bullet Point 1 RU");
    IF newBulletPoint1RU = '' THEN EXIT;
    TempBlob.Blob:="Bullet Point 1 RU";
    TempBlob.WriteAsText(newBulletPoint1RU, TEXTENCODING::UTF8);
    "Bullet Point 1 RU":=TempBlob.Blob;
    if MODIFY then;
  end;
  procedure GetBulletPoint1RU(): Text;
  begin
    EXIT(GetBulletPoint1RUCalculated);
  end;
  local procedure GetBulletPoint1RUCalculated(): Text;
  var TempBlob: Record TempBlob temporary;
  CR: text[1];
  begin
    IF NOT "Bullet Point 1 RU".HASVALUE THEN EXIT('');
    CALCFIELDS("Bullet Point 1 RU");
    CR[1]:=10;
    TempBlob.Blob:="Bullet Point 1 RU";
    EXIT(TempBlob.ReadAsText(CR, TEXTENCODING::UTF8));
  end;
  // BulletPoint2RU
  procedure SetBulletPoint2RU(newBulletPoint2RU: Text)var TempBlob: Record TempBlob temporary;
  begin
    CLEAR("Bullet Point 2 RU");
    IF newBulletPoint2RU = '' THEN EXIT;
    TempBlob.Blob:="Bullet Point 2 RU";
    TempBlob.WriteAsText(newBulletPoint2RU, TEXTENCODING::UTF8);
    "Bullet Point 2 RU":=TempBlob.Blob;
    if MODIFY then;
  end;
  procedure GetBulletPoint2RU(): Text;
  begin
    EXIT(GetBulletPoint2RUCalculated);
  end;
  local procedure GetBulletPoint2RUCalculated(): Text;
  var TempBlob: Record TempBlob temporary;
  CR: text[1];
  begin
    IF NOT "Bullet Point 2 RU".HASVALUE THEN EXIT('');
    CALCFIELDS("Bullet Point 2 RU");
    CR[1]:=10;
    TempBlob.Blob:="Bullet Point 2 RU";
    EXIT(TempBlob.ReadAsText(CR, TEXTENCODING::UTF8));
  end;
  // BulletPoint3RU
  procedure SetBulletPoint3RU(newBulletPoint3RU: Text)var TempBlob: Record TempBlob temporary;
  begin
    CLEAR("Bullet Point 3 RU");
    IF newBulletPoint3RU = '' THEN EXIT;
    TempBlob.Blob:="Bullet Point 3 RU";
    TempBlob.WriteAsText(newBulletPoint3RU, TEXTENCODING::UTF8);
    "Bullet Point 3 RU":=TempBlob.Blob;
    if MODIFY then;
  end;
  procedure GetBulletPoint3RU(): Text;
  begin
    EXIT(GetBulletPoint3RUCalculated);
  end;
  local procedure GetBulletPoint3RUCalculated(): Text;
  var TempBlob: Record TempBlob temporary;
  CR: text[1];
  begin
    IF NOT "Bullet Point 3 RU".HASVALUE THEN EXIT('');
    CALCFIELDS("Bullet Point 3 RU");
    CR[1]:=10;
    TempBlob.Blob:="Bullet Point 3 RU";
    EXIT(TempBlob.ReadAsText(CR, TEXTENCODING::UTF8));
  end;
  // BulletPoint4RU
  procedure SetBulletPoint4RU(newBulletPoint4RU: Text)var TempBlob: Record TempBlob temporary;
  begin
    CLEAR("Bullet Point 4 RU");
    IF newBulletPoint4RU = '' THEN EXIT;
    TempBlob.Blob:="Bullet Point 4 RU";
    TempBlob.WriteAsText(newBulletPoint4RU, TEXTENCODING::UTF8);
    "Bullet Point 4 RU":=TempBlob.Blob;
    if MODIFY then;
  end;
  procedure GetBulletPoint4RU(): Text;
  begin
    EXIT(GetBulletPoint4RUCalculated);
  end;
  local procedure GetBulletPoint4RUCalculated(): Text;
  var TempBlob: Record TempBlob temporary;
  CR: text[1];
  begin
    IF NOT "Bullet Point 4 RU".HASVALUE THEN EXIT('');
    CALCFIELDS("Bullet Point 4 RU");
    CR[1]:=10;
    TempBlob.Blob:="Bullet Point 4 RU";
    EXIT(TempBlob.ReadAsText(CR, TEXTENCODING::UTF8));
  end;
  // IngredientsRU
  procedure SetIngredientsRU(newIngredientsRU: Text)var TempBlob: Record TempBlob temporary;
  begin
    CLEAR("Ingredients RU");
    IF newIngredientsRU = '' THEN EXIT;
    TempBlob.Blob:="Indications RU";
    TempBlob.WriteAsText(newIngredientsRU, TEXTENCODING::UTF8);
    "Ingredients RU":=TempBlob.Blob;
    if MODIFY then;
  end;
  procedure GetIngredientsRU(): Text;
  begin
    EXIT(GetIngredientsRUCalculated);
  end;
  local procedure GetIngredientsRUCalculated(): Text;
  var TempBlob: Record TempBlob temporary;
  CR: text[1];
  begin
    IF NOT "Ingredients RU".HASVALUE THEN EXIT('');
    CALCFIELDS("Ingredients RU");
    CR[1]:=10;
    TempBlob.Blob:="Ingredients RU";
    EXIT(TempBlob.ReadAsText(CR, TEXTENCODING::UTF8));
  end;
  // IndicationsRU
  procedure SetIndicationsRU(newIndicationsRU: Text)var TempBlob: Record TempBlob temporary;
  begin
    CLEAR("Indications RU");
    IF newIndicationsRU = '' THEN EXIT;
    TempBlob.Blob:="Indications RU";
    TempBlob.WriteAsText(newIndicationsRU, TEXTENCODING::UTF8);
    "Indications RU":=TempBlob.Blob;
    if MODIFY then;
  end;
  procedure GetIndicationsRU(): Text;
  begin
    EXIT(GetIndicationsRUCalculated);
  end;
  local procedure GetIndicationsRUCalculated(): Text;
  var TempBlob: Record TempBlob temporary;
  CR: text[1];
  begin
    IF NOT "Indications RU".HASVALUE THEN EXIT('');
    CALCFIELDS("Indications RU");
    CR[1]:=10;
    TempBlob.Blob:="Indications RU";
    EXIT(TempBlob.ReadAsText(CR, TEXTENCODING::UTF8));
  end;
  // DirectionsRU
  procedure SetDirectionsRU(newDirectionsRU: Text)var TempBlob: Record TempBlob temporary;
  begin
    CLEAR("Directions RU");
    IF newDirectionsRU = '' THEN EXIT;
    TempBlob.Blob:="Directions RU";
    TempBlob.WriteAsText(newDirectionsRU, TEXTENCODING::UTF8);
    "Directions RU":=TempBlob.Blob;
    if MODIFY then;
  end;
  procedure GetDirectionsRU(): Text;
  begin
    EXIT(GetDirectionsRUCalculated);
  end;
  local procedure GetDirectionsRUCalculated(): Text;
  var TempBlob: Record TempBlob temporary;
  CR: text[1];
  begin
    IF NOT "Directions RU".HASVALUE THEN EXIT('');
    CALCFIELDS("Directions RU");
    CR[1]:=10;
    TempBlob.Blob:="Directions RU";
    EXIT(TempBlob.ReadAsText(CR, TEXTENCODING::UTF8));
  end;
}
