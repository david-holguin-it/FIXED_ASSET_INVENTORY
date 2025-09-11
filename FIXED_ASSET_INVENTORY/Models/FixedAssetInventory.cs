namespace FIXED_ASSET_INVENTORY.Models
{
    public class FixedAssetItem
    {
        public int id { get; set;}
        public string manufacturerName { get; set; }
        public string partyManufacturerName { get; set; }
        public string materialNumber { get; set; }
        public string productName { get; set; }
        public string description { get; set; }
        public int quantity { get; set; }
        public float unitPrice { get; set; }
        public float totalPrice { get; set; }
        public float unitPriceUSD { get; set; }
        public float totalUSD { get; set; }
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
        public string NOTE { get; set; }

    };

}
