namespace FIXED_ASSET_INVENTORY.Models
{
    public class FixedAssetItem
    {
        public int id { get; set;}
        public string manufacturerName { get; set; }
        public string partyManufacturerName { get; set; }   // TBD se elimina?
        public string materialNumber { get; set; }
        public string productName { get; set; }
        public string description { get; set; } 
        public float purchaseValue { get; set; }
    //    public float totalPrice { get; set; }       //
     //   public float unitPriceUSD { get; set; }
     //   public float totalUSD { get; set; }         //
        public string paymentTerms { get; set; }
        public string purchaseOrderNo { get; set; }
        public string contractNo { get; set; }
        public string signOff { get; set; }
        public string remark { get; set; }
        public string materialsSent { get; set;}
        public string department { get; set; }
        public string manager { get; set; }
        public string fixedAssetNumber { get; set; }
        public string serialNumber { get; set; }
        public string location { get; set; }
        public string PIC { get; set; } 
        public float accumulatedDepreciation { get; set; } 
        public float netBookValue { get; set; }
        public int usefulLife { get; set; }
        public DateTime capitalizationDate { get; set; } 
        public string updatedBy { get; set; }
    };
}
