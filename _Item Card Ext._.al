pageextension 50001 "Item Card Ext." extends "Item Card"
{
  layout
  {
    // Add changes to page layout here
    addlast(Item)
    {
      field("No. 2";"No. 2")
      {
        ApplicationArea = All;
        Importance = Additional;
      }
    }
    addafter(Control1907509201)
    {
      group(groupItemDescription)
      {
        CaptionML = ENU='Item Descriptions';
        Editable = EditAllowed;

        field("Manufacturer Code";"Manufacturer Code")
        {
          ApplicationArea = All;
        }
        field("Brand Code";"Brand Code")
        {
          ApplicationArea = All;
        }
        field("Name RU";ItemDescription."Name RU")
        {
          ApplicationArea = All;

          trigger OnValidate()begin
            ItemDescriptionModify;
          end;
        }
        field("Name RU 2";ItemDescription."Name RU 2")
        {
          ApplicationArea = All;

          trigger OnValidate()begin
            ItemDescriptionModify;
          end;
        }
        group(groupDescription)
        {
          CaptionML = ENU='Description';

          field("iDescription";txtDescription)
          {
            ApplicationArea = All;
            ShowCaption = false;
            MultiLine = true;

            trigger OnValidate()begin
              if not ItemDescriptionExist then ItemDescription.InitItemDescription("No.");
              ItemDescription.SetDescription(txtDescription);
            end;
          }
        }
        group(groupIngredients)
        {
          CaptionML = ENU='Ingredients';

          field("Ingredients";txtIngredients)
          {
            ApplicationArea = All;
            ShowCaption = false;
            MultiLine = true;

            trigger OnValidate()begin
              if not ItemDescriptionExist then ItemDescription.InitItemDescription("No.");
              ItemDescription.SetIngredients(txtIngredients);
            end;
          }
        }
        group(groupIndications)
        {
          CaptionML = ENU='Indications';

          field("Indications";txtIndications)
          {
            ApplicationArea = All;
            ShowCaption = false;
            MultiLine = true;

            trigger OnValidate()begin
              if not ItemDescriptionExist then ItemDescription.InitItemDescription("No.");
              ItemDescription.SetIndications(txtIndications);
            end;
          }
        }
        group(groupDirections)
        {
          CaptionML = ENU='Directions';

          field("Directions";txtDirections)
          {
            ApplicationArea = All;
            ShowCaption = false;
            MultiLine = true;

            trigger OnValidate()begin
              if not ItemDescriptionExist then ItemDescription.InitItemDescription("No.");
              ItemDescription.SetDirections(txtDirections);
            end;
          }
        }
        group(groupDescriptionRU)
        {
          CaptionML = ENU='Description RU';

          field("Description RU";txtDescriptionRU)
          {
            ApplicationArea = All;
            ShowCaption = false;
            MultiLine = true;

            trigger OnValidate()begin
              if not ItemDescriptionExist then ItemDescription.InitItemDescription("No.");
              ItemDescription.SetDescriptionRU(txtDescriptionRU);
            end;
          }
        }
        group(groupIngredientsRU)
        {
          CaptionML = ENU='Ingredients RU';

          field("Ingredients RU";txtIngredientsRU)
          {
            ApplicationArea = All;
            ShowCaption = false;
            MultiLine = true;

            trigger OnValidate()begin
              if not ItemDescriptionExist then ItemDescription.InitItemDescription("No.");
              ItemDescription.SetIngredientsRU(txtIngredientsRU);
            end;
          }
        }
        group(groupIndicationsRU)
        {
          CaptionML = ENU='Indications RU';

          field("Indications RU";txtIndicationsRU)
          {
            ApplicationArea = All;
            ShowCaption = false;
            MultiLine = true;

            trigger OnValidate()begin
              if not ItemDescriptionExist then ItemDescription.InitItemDescription("No.");
              ItemDescription.SetIndicationsRU(txtIndicationsRU);
            end;
          }
        }
        group(groupDirectionsRU)
        {
          CaptionML = ENU='Directions RU';

          field("Directions RU";txtDirectionsRU)
          {
            ApplicationArea = All;
            ShowCaption = false;
            MultiLine = true;

            trigger OnValidate()begin
              if not ItemDescriptionExist then ItemDescription.InitItemDescription("No.");
              ItemDescription.SetDirectionsRU(txtDirectionsRU);
            end;
          }
        }
        group(groupWarning)
        {
          CaptionML = ENU='Warning';

          field("Warning";txtWarning)
          {
            ApplicationArea = All;
            ShowCaption = false;
            MultiLine = true;

            trigger OnValidate()begin
              if not ItemDescriptionExist then ItemDescription.InitItemDescription("No.");
              ItemDescription.SetWarning(txtWarning);
            end;
          }
        }
      }
      group(BulletPointArea)
      {
        CaptionML = ENU='Bullet Points';
        Editable = EditAllowed;

        group(groupBulletPoint1)
        {
          CaptionML = ENU='Bullet Point 1';

          field("BulletPoint1";txtBulletPoint1)
          {
            ApplicationArea = All;
            ShowCaption = false;
            MultiLine = true;

            trigger OnValidate()begin
              if not ItemDescriptionExist then ItemDescription.InitItemDescription("No.");
              ItemDescription.SetBulletPoint1(txtBulletPoint1);
            end;
          }
        }
        group(groupBulletPoint2)
        {
          CaptionML = ENU='Bullet Point 2';

          field("BulletPoint2";txtBulletPoint2)
          {
            ApplicationArea = All;
            ShowCaption = false;
            MultiLine = true;

            trigger OnValidate()begin
              if not ItemDescriptionExist then ItemDescription.InitItemDescription("No.");
              ItemDescription.SetBulletPoint2(txtBulletPoint2);
            end;
          }
        }
        group(groupBulletPoint3)
        {
          CaptionML = ENU='Bullet Point 3';

          field("BulletPoint3";txtBulletPoint3)
          {
            ApplicationArea = All;
            ShowCaption = false;
            MultiLine = true;

            trigger OnValidate()begin
              if not ItemDescriptionExist then ItemDescription.InitItemDescription("No.");
              ItemDescription.SetBulletPoint3(txtBulletPoint3);
            end;
          }
        }
        group(groupBulletPoint4)
        {
          CaptionML = ENU='Bullet Point 4';

          field("BulletPoint4";txtBulletPoint4)
          {
            ApplicationArea = All;
            ShowCaption = false;
            MultiLine = true;

            trigger OnValidate()begin
              if not ItemDescriptionExist then ItemDescription.InitItemDescription("No.");
              ItemDescription.SetBulletPoint4(txtBulletPoint4);
            end;
          }
        }
        field("Bullet Point 5";ItemDescription."Bullet Point 5")
        {
          ApplicationArea = All;

          trigger OnValidate()begin
            ItemDescriptionModify;
          end;
        }
        group(groupBulletPoint1RU)
        {
          CaptionML = ENU='Bullet Point 1 RU';

          field("Bullet Point 1 RU";txtBulletPoint1RU)
          {
            ApplicationArea = All;
            ShowCaption = false;
            MultiLine = true;

            trigger OnValidate()begin
              if not ItemDescriptionExist then ItemDescription.InitItemDescription("No.");
              ItemDescription.SetBulletPoint1RU(txtBulletPoint1RU);
            end;
          }
        }
        group(groupBulletPoint2RU)
        {
          CaptionML = ENU='Bullet Point 2 RU';

          field("Bullet Point 2 RU";txtBulletPoint2RU)
          {
            ApplicationArea = All;
            ShowCaption = false;
            MultiLine = true;

            trigger OnValidate()begin
              if not ItemDescriptionExist then ItemDescription.InitItemDescription("No.");
              ItemDescription.SetBulletPoint2RU(txtBulletPoint2RU);
            end;
          }
        }
        group(groupBulletPoint3RU)
        {
          CaptionML = ENU='Bullet Point 3 RU';

          field("Bullet Point 3 RU";txtBulletPoint3RU)
          {
            ApplicationArea = All;
            ShowCaption = false;
            MultiLine = true;

            trigger OnValidate()begin
              if not ItemDescriptionExist then ItemDescription.InitItemDescription("No.");
              ItemDescription.SetBulletPoint3RU(txtBulletPoint3RU);
            end;
          }
        }
        group(groupBulletPoint4RU)
        {
          CaptionML = ENU='Bullet Point 4 RU';

          field("Bullet Point 4 RU";txtBulletPoint4RU)
          {
            ApplicationArea = All;
            ShowCaption = false;
            MultiLine = true;

            trigger OnValidate()begin
              if not ItemDescriptionExist then ItemDescription.InitItemDescription("No.");
              ItemDescription.SetBulletPoint4RU(txtBulletPoint4RU);
            end;
          }
        }
        field("Bullet Point 5 RU";ItemDescription."Bullet Point 5 RU")
        {
          ApplicationArea = All;

          trigger OnValidate()begin
            ItemDescriptionModify;
          end;
        }
      }
      group(SearchArea)
      {
        CaptionML = ENU='Searches';
        Editable = EditAllowed;

        group(groupSearchTerms)
        {
          CaptionML = ENU='Search Terms';

          field("Search Terms";txtSearchTerms)
          {
            ApplicationArea = All;
            ShowCaption = false;
            MultiLine = true;

            trigger OnValidate()begin
              if not ItemDescriptionExist then ItemDescription.InitItemDescription("No.");
              ItemDescription.SetSearchTerms(txtSearchTerms);
            end;
          }
        }
        group(groupSearchTermsForGoogleOnly)
        {
          CaptionML = ENU='Search Terms For Google Only';

          field("Search Terms For Google Only";txtSearchTermsForGoogleOnly)
          {
            ApplicationArea = All;
            ShowCaption = false;
            MultiLine = true;

            trigger OnValidate()begin
              if not ItemDescriptionExist then ItemDescription.InitItemDescription("No.");
              ItemDescription.SetSearchTermsForGoogleOnly(txtSearchTermsForGoogleOnly);
            end;
          }
        }
      }
      group(groupURLArea)
      {
        CaptionML = ENU='URLs';
        Editable = EditAllowed;

        field("Main Image URL";ItemDescription."Main Image URL")
        {
          ApplicationArea = All;

          trigger OnValidate()begin
            ItemDescriptionModify;
          end;
        }
        field("Other Image URL";ItemDescription."Other Image URL")
        {
          ApplicationArea = All;

          trigger OnValidate()begin
            ItemDescriptionModify;
          end;
        }
        field("Label Image URL";ItemDescription."Label Image URL")
        {
          ApplicationArea = All;

          trigger OnValidate()begin
            ItemDescriptionModify;
          end;
        }
        field("Label Image URL 2";ItemDescription."Label Image URL 2")
        {
          ApplicationArea = All;

          trigger OnValidate()begin
            ItemDescriptionModify;
          end;
        }
      }
      group(groupWebArea)
      {
        CaptionML = ENU='Web';
        Editable = EditAllowed;

        field("New Item to";ItemDescription.New)
        {
          ApplicationArea = All;

          trigger OnValidate()begin
            ItemDescriptionModify;
          end;
        }
        field("Sell-out to";ItemDescription."Sell-out")
        {
          ApplicationArea = All;

          trigger OnValidate()begin
            ItemDescriptionModify;
          end;
        }
        field(Barcode;ItemDescription.Barcode)
        {
          ApplicationArea = All;

          trigger OnValidate()begin
            ItemDescriptionModify;
          end;
        }
      }
    }
  }
  actions
  {
  }
  var ItemDescriptionExist: Boolean;
  EditAllowed: Boolean;
  ItemDescription: Record "Item Description";
  txtWarning: Text;
  txtDescription: Text;
  txtBulletPoint1: Text;
  txtBulletPoint2: Text;
  txtBulletPoint3: Text;
  txtBulletPoint4: Text;
  txtSearchTerms: Text;
  txtSearchTermsForGoogleOnly: Text;
  txtIngredients: Text;
  txtIndications: Text;
  txtDirections: Text;
  txtDescriptionRU: Text;
  txtBulletPoint1RU: Text;
  txtBulletPoint2RU: Text;
  txtBulletPoint3RU: Text;
  txtBulletPoint4RU: Text;
  txtIngredientsRU: Text;
  txtIndicationsRU: Text;
  txtDirectionsRU: Text;
  trigger OnAfterGetRecord()begin
    Clear(txtWarning);
    Clear(txtDescription);
    Clear(txtBulletPoint1);
    Clear(txtBulletPoint2);
    Clear(txtBulletPoint3);
    Clear(txtBulletPoint4);
    Clear(txtSearchTerms);
    Clear(txtSearchTermsForGoogleOnly);
    Clear(txtIngredients);
    Clear(txtIndications);
    Clear(txtDirections);
    Clear(txtDescriptionRU);
    Clear(txtBulletPoint1RU);
    Clear(txtBulletPoint2RU);
    Clear(txtBulletPoint3RU);
    Clear(txtBulletPoint4RU);
    Clear(txtIngredientsRU);
    Clear(txtIndicationsRU);
    Clear(txtDirectionsRU);
    EditAllowed:=CurrPage.EDITABLE;
    ItemDescriptionExist:=false;
    if(xRec."No." <> "No.")then if ItemDescription.Get("No.")then begin
        ItemDescriptionExist:=true;
        txtWarning:=ItemDescription.GetWarning;
        txtDescription:=ItemDescription.GetDescription;
        txtBulletPoint1:=ItemDescription.GetBulletPoint1;
        txtBulletPoint2:=ItemDescription.GetBulletPoint2;
        txtBulletPoint3:=ItemDescription.GetBulletPoint3;
        txtBulletPoint4:=ItemDescription.GetBulletPoint4;
        txtSearchTerms:=ItemDescription.GetSearchTerms;
        txtSearchTermsForGoogleOnly:=ItemDescription.GetSearchTermsForGoogleOnly;
        txtIngredients:=ItemDescription.GetIngredients;
        txtIndications:=ItemDescription.GetIndications;
        txtDirections:=ItemDescription.GetDirections;
        txtDescriptionRU:=ItemDescription.GetDescriptionRU;
        txtBulletPoint1RU:=ItemDescription.GetBulletPoint1RU;
        txtBulletPoint2RU:=ItemDescription.GetBulletPoint2RU;
        txtBulletPoint3RU:=ItemDescription.GetBulletPoint3RU;
        txtBulletPoint4RU:=ItemDescription.GetBulletPoint4RU;
        txtIngredientsRU:=ItemDescription.GetIngredientsRU;
        txtIndicationsRU:=ItemDescription.GetIndicationsRU;
        txtDirectionsRU:=ItemDescription.GetDirectionsRU;
      end;
  end;
  local procedure ItemDescriptionModify()begin
    if not ItemDescriptionExist then ItemDescription.InitItemDescription("No.");
    ItemDescription.Modify();
  end;
}
