table 50001 Brand
{
  DataClassification = ToBeClassified;

  fields
  {
    field(1;"Code";Code[20])
    {
      DataClassification = ToBeClassified;
    }
    field(2;"Manufacturer Code";Code[10])
    {
      DataClassification = ToBeClassified;
    }
    field(3;Name;Text[50])
    {
      DataClassification = ToBeClassified;
    }
  }
  keys
  {
    key(PK;"Code", "Manufacturer Code")
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
}
