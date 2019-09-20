tableextension 50000 "Manufacturer Ext." extends Manufacturer
{
  fields
  {
    // Add changes to table fields here
    field(50000;Address;Text[150])
    {
      DataClassification = ToBeClassified;
    }
    field(50001;"Name RU";Text[50])
    {
      DataClassification = ToBeClassified;
    }
  }
}
