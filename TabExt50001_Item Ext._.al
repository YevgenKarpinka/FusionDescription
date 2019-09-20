tableextension 50001 "Item Ext." extends Item
{
  fields
  {
    // Add changes to table fields here
    modify("Manufacturer Code")
    {
    trigger OnAfterValidate()begin
      if xRec."Manufacturer Code" <> "Manufacturer Code" then "Brand Code":='';
    end;
    }
    field(50001;"Brand Code";Code[20])
    {
      DataClassification = ToBeClassified;
      TableRelation = Brand;

      trigger OnValidate()var Brand: Record Brand;
      begin
        if xRec."Brand Code" <> "Brand Code" then begin
          Brand.Reset();
          Brand.SetRange(Code, "Brand Code");
          if Brand.FindFirst()then "Manufacturer Code":=Brand."Manufacturer Code";
        end;
      end;
    }
  }
  var myInt: Integer;
}
