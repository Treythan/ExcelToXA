using ClosedXML.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net;
using System.Text;
using System.Xml;

namespace ExcelToXA
{
    public partial class Main : Form
    {
        string excelFilePath;
        int excelViewIndex;
        int poViewIndex;
        List<Order> orders = new List<Order>();
        public UserSettings userSettings = new UserSettings();

        public Main()
        {
            InitializeComponent();
            Program.main = this;
            Program.config = new Configuration();
            Program.config.ShowDialog();
            uploadBtn.Enabled = false;
            delItem.Enabled = false;
            refreshTT.SetToolTip(refreshBtn, "Refresh Purchase Orders");
        }
        private void LoadDataToPreview()
        {
            excelDataPreview.Items.Clear();
            XLWorkbook wb = new XLWorkbook(excelFilePath);
            IXLWorksheet ws = wb.Worksheets.FirstOrDefault();
            excelDataPreview.Columns.Add(new ColumnHeader { Width = 105, Text = "Part Number" });
            excelDataPreview.Columns.Add(new ColumnHeader { Width = 105, Text = "Quantity" });

            Vendor vendor = userSettings.Vendor;

            foreach (IXLRow row in ws.RowsUsed(XLCellsUsedOptions.AllContents))
            {
                if (vendor.HasHeaders && row.RowNumber() == 1) continue;
                if (row.Cell(vendor.PNColumn).Value.ToString() == "" || row.Cell(vendor.QTYColumn).Value.ToString() == "") continue;
                var item = new ListViewItem { Text = row.Cell(vendor.PNColumn).Value.ToString() };
                item.SubItems.Add(new ListViewItem.ListViewSubItem { Text = row.Cell(vendor.QTYColumn).Value.ToString() });
                item.SubItems.Add(new ListViewItem.ListViewSubItem { Text = row.Cell(vendor.QTYColumn).Value.ToString() });
                excelDataPreview.Items.Add(item);
            }
        }

        public void BeginItemsUpload()
        {
            string PO = poPreview.Items[poViewIndex].Text;

            Order order = orders.FirstOrDefault(x => x.OrderNumber == PO);
            if (order != null)
            {
                progressBar1.Maximum = excelDataPreview.Items.Count;
                foreach (ListViewItem item in excelDataPreview.Items)
                {
                    ItemDetails details = GetItemDetails(item.Text, order.Details.Warehouse, item.SubItems[1].Text);
                    UploadItem(order.OrderNumber, details);
                    progressBar1.Value++;
                }
                MessageBox.Show("Upload Complete!", "Excel PO Loading", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        public ItemDetails GetItemDetails(string itemNumber, string warehouse, string qtyRequested)
        {
            string req = $"<?xml version='1.0' encoding='UTF-8'?><!DOCTYPE System-Link SYSTEM 'SystemLinkRequest.dtd'>"+
                $"<System-Link><Login userId='{Program.xaLogin.Key}' password='{Program.xaLogin.Value}' maxIdle='900000'\r\n"+
                $"          properties='com.pjx.cas.domain.EnvironmentId={userSettings.Environment.DisplayText.Substring(0, 2)},\r\n" +
                "               com.pjx.cas.domain.SystemName=INFXAPRD,\r\n" +
                "               com.pjx.cas.user.LanguageId=en'/><Request sessionHandle='*current' workHandle='*new'\r\n" +
                "          broker='EJB' maxIdle='1000'><QueryObject name='queryObject_ItemWarehouse_DefaultR9' " +
                "               domainClass='com.mapics.epdm.ItemWarehouse' includeMetaData='true'><Pql><![CDATA[SELECT relatedItemRevision.description,\r\n" +
                "                          activeRecordCode,planner,relatedItemRevision.itemType,\r\n" +
                "                          itemClass,relatedItemRevision.itemClass,\r\n" +
                "                          itemAccountingClass,leadTimeCode,relatedItemWarehouseExtension.defaultInspectionLocation,\r\n" +
                "                          defaultStockLocation,relatedItemWarehouseExtension.defaultInTransitWarehouse,\r\n" +
                "                          relatedItemWarehouseExtension.defaultInTransitLocation,\r\n" +
                "                          relatedItemRevision.stockingUm,floorStockCode,\r\n" +
                "                          userFieldSwitchA,stockConfigurations,\r\n" +
                "                          relatedItemPlan.itemRescheduleCode,\r\n" +
                "                          lastMaintainedDate,relatedItemWarehouseExtension.receiptsTolerancePercentage,\r\n" +
                "                          relatedZ1783EF80069470220237AAAAAAAADKL.ixcrte,\r\n" +
                "                          relatedZ1783EF80069470220237AAAAAAAADKL.ixmetd,\r\n" +
                "                          userFieldCodeB,quantityOnHand,totalQuantityAllocated,\r\n" +
                "                          totalQuantityAvailableOnHand,safetyStock,\r\n" +
                "                          totalQuantityPendingAllocation,totalQuantityOnOrder,\r\n" +
                "                          totalQuantityAvailable,relatedItemWarehouseExtension.inventoryStatus,\r\n" +
                "                          relatedItemWarehouseExtension.externallyControlledQuantity,\r\n" +
                "                          quantityAllocatedCustomerOrders,quantityAllocatedProduction,\r\n" +
                "                          totalQuantityPendingAllocation,quantityOnOrderProduction,\r\n" +
                "                          quantityOnOrderPurchase,inTransitOnHandQuantity,\r\n" +
                "                          inTransitOnHandQuantity,relatedItemWarehouseExtension.incomingInTransitTotalQuantity,\r\n" +
                "                          relatedItemWarehouseExtension.onHoldQuantityPlanning,\r\n" +
                "                          beginningInventory,estimateAnnualUsage,\r\n" +
                "                          averageEndingInventory,averageTurnover,\r\n" +
                "                          quantityIssuedThisPeriod,quantityIssuedThisYear,\r\n" +
                "                          dateOfLastIssue,quantityReceivedThisPeriod,\r\n" +
                "                          quantityAdjustedThisPeriod,quantityUsedThisPeriod,\r\n" +
                "                          quantityUsedThisYear,dateOfLastUsage,\r\n" +
                "                          quantityScrappedThisPeriod,quantityScrappedThisYear,\r\n" +
                "                          dateOfLastScrap,quantityOnHand,standardUnitCost,\r\n" +
                "                          averageUnitCost,lastUnitCost,averageSalesPerPeriod,\r\n" +
                "                          dateOfLastSale,quantitySoldThisPeriod,\r\n" +
                "                          quantitySoldThisYear,salesAmountThisPeriod,\r\n" +
                "                          salesAmountThisYear,salesCostThisPeriod,\r\n" +
                "                          salesCostThisYear,usageCostThisPeriod,\r\n" +
                "                          usageCostThisYear,standardUnitCost,\r\n" +
                "                          averageUnitCost,lastUnitCost,relatedItemRevision.unitCostDefault,\r\n" +
                "                          relatedItemRevision.relatedItemRevisionCost.standardUnitCost,\r\n" +
                "                          relatedItemRevision.relatedItemRevisionCost.currentUnitCost,\r\n" +
                "                          relatedItemPlan.orderPolicyCode,relatedItemPlan.reschedulingFrozenZone,\r\n" +
                "                          relatedItemPlan.daysSupplyPerOrder,\r\n" +
                "                          relatedItemPlan.shrinkageFactor,relatedItemPlan.minimumQuantity,\r\n" +
                "                          relatedItemPlan.maximumQuantity,relatedItemPlan.multipleQuantity,\r\n" +
                "                          fixedOrderQuantity,userFieldQuantity1,\r\n" +
                "                          orderPoint,relatedItemPlan.manufactureAutoReleaseCode,\r\n" +
                "                          smoothingCode,smoothingStartDate,relatedItemPlan.minimumDaysToReschedule,\r\n" +
                "                          relatedItemPlan.masterLevelItemCode,\r\n" +
                "                          relatedItemPlan.masterLevelForecastCode,\r\n" +
                "                          forecastQuantityPeriod,relatedItemPlan.numberOfForecastPeriods,\r\n" +
                "                          relatedItemPlan.daysForecastPeriod,\r\n" +
                "                          relatedItemPlan.forecastCode,relatedItemPlan.combineRequirements,\r\n" +
                "                          relatedItemPlan.planExpectedOrdersCode,\r\n" +
                "                          relatedItemPlan.planCustomerOrdersCode,\r\n" +
                "                          relatedItemPlan.contractRequiredCode,\r\n" +
                "                          relatedItemPlan.priceBreakConversionFactor,\r\n" +
                "                          relatedItemPlan.periodIntervalCode,\r\n" +
                "                          relatedItemPlan.masterLevelPrintCode,\r\n" +
                "                          relatedItemPlan.maximumNumberOfLinesItem,\r\n" +
                "                          relatedItemPlan.multiSourceCode,includeInventoryBalance,\r\n" +
                "                          replanFlag,leadTimeCode,leadTimeManufacturing,\r\n" +
                "                          leadTimeVariableManufacturing,leadTimeAdjustmentManufacturing,\r\n" +
                "                          leadTimeAverageManufacturing,leadTimeCumulativeManufacturing,\r\n" +
                "                          leadTimeCumulativeMaterial,relatedItemRevision.relatedItemRevisionCost.standardLotSize,\r\n" +
                "                          leadTimeCode,leadTimePurchase,leadTimeAdjustmentPurchase,\r\n" +
                "                          leadTimeAveragePurchase,leadTimeReview,\r\n" +
                "                          leadTimeVendor,leadTimeSafety,leadTimeCumulativeMaterial,\r\n" +
                "                          relatedItemRevision.relatedItemRevisionCost.standardLotSize,\r\n" +
                "                          primaryVendor,purchaseUm,umConversionFactor,\r\n" +
                "                          relatedItemPlan.purchaseAutoReleaseCode,\r\n" +
                "                          relatedItemPlan.firmTimeFence,relatedItemPlan.authorizedTimeFence,\r\n" +
                "                          relatedItemPlan.purchasePlanningScheduleProfile,\r\n" +
                "                          relatedItemPlan.masterScheduleItemCode,\r\n" +
                "                          relatedItemPlan.productionFamilyPlanner,\r\n" +
                "                          relatedItemPlan.mpsPlanningSourceCode,\r\n" +
                "                          relatedItemPlan.demandTimeFence,relatedItemPlan.resourceNumber,\r\n" +
                "                          relatedItemPlan.buildResourceProfiles,\r\n" +
                "                          relatedItemPlan.plannedBy,relatedItemPlan.constrainedPart,\r\n" +
                "                          userFieldCodeC,cycleCountCode,userFieldCodeA,\r\n" +
                "                          cycleCountClass,cycleCountActivityFlag,\r\n" +
                "                          quantityOnHandAtCount,transactionCompare,\r\n" +
                "                          transactionActivityCount,forceCycleCountFlag,\r\n" +
                "                          nextCountDate,lastCountDate,countAccuracyTolerancePercent,\r\n" +
                "                          scheduleControlCode,extractSourceCode,\r\n" +
                "                          carryForwardCode,smoothingCode,smoothingStartDate,\r\n" +
                "                          primaryProductionLine,scheduleGroup,\r\n" +
                "                          containerDescription,quantityPerContainer,\r\n" +
                "                          lotSizing,defaultStockLocation,backflushMethod,\r\n" +
                "                          relatedItemRevision.stockingUm,totalQuantityAvailable,\r\n" +
                "                          quantityOnHand,warehouse,totalQuantityAllocated,\r\n" +
                "                          totalQuantityOnOrder,item REQUIRED \r\n" +
                "                          relatedItemPlan,relatedItemWarehouseExtension \r\n" +
                "                          WHERE warehouse = '{warehouse}' AND item = '{itemNumber}']]></Pql></QueryObject></Request></System-Link>\r\n";
            HttpWebResponse response = SendWebRequest(req);
            if (response.StatusCode == HttpStatusCode.OK)
            {
                Stream responseStream = response.GetResponseStream();
                string responseStr = new StreamReader(responseStream).ReadToEnd();
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(responseStr);
                string json = JsonConvert.SerializeXmlNode(doc, Newtonsoft.Json.Formatting.Indented);
                JObject resp = JObject.Parse(json);
                JArray objects = (JArray)JsonConvert.DeserializeObject(resp["System-Link"]["Response"]["QueryObjectResponse"]["DomainEntity"]["Property"].ToString());
                ItemDetails details = new ItemDetails();
                details.OrderQuantityRequested = qtyRequested;
                foreach (JObject property in objects)
                {
                    switch (property["@path"].ToString())
                    {
                        case "item":
                            details.ItemNumber = GetValue(property);
                            break;
                        case "relatedItemRevision.relatedItemRevisionCost.currentUnitCost":
                            details.UnitPrice = GetValue(property);
                            break;
                        case "warehouse":
                            details.Warehouse = GetValue(property);
                            break;
                        case "planner":
                            details.Planner = GetValue(property);
                            break;
                        case "purchaseUm":
                            details.UnitOfMeasure = GetValue(property);
                            break;
                        case "relatedItemRevision.description":
                            details.ItemDescription = GetValue(property);
                            break;
                        case "leadTimeVendor":
                            details.VendorLeadTime = GetValue(property);
                            break;
                    }
                }
                return details;
            }
            else return null;
        }

        public void UploadItem(string poNumber, ItemDetails id)
        {
            string req = $"<?xml version='1.0' encoding='UTF-8'?><!DOCTYPE System-Link SYSTEM 'SystemLinkRequest.dtd'><System-Link>" +
                $"<Login userId='{Program.xaLogin.Key}' password='{Program.xaLogin.Value}' maxIdle='900000'\r\n" +
                $"          properties='com.pjx.cas.domain.EnvironmentId={userSettings.Environment.DisplayText.Substring(0, 2)},\r\n" +
                $"               com.pjx.cas.domain.SystemName=INFXAPRD,\r\n" +
                $"               com.pjx.cas.user.LanguageId=en'/><Request sessionHandle='*current' workHandle='*new'\r\n" +
                $"          broker='EJB' maxIdle='1000'><Create name='createObject_PurchaseOrderItem' domainClass='com.mapics.pm.PoItem'>" +
                $"<ApplyTemplate><![CDATA[(none)]]></ApplyTemplate><MaintenanceOptions><Option optionName='includeVendorItemComments'>" +
                $"<Value><![CDATA[1]]></Value></Option><Option optionName='includeItemReceivingOperations'><Value><![CDATA[0]]></Value></Option></MaintenanceOptions>" +
                $"<DomainEntity><Key><Property path='order'><Value><![CDATA[{poNumber}]]></Value></Property></Key><Property path='contract'>" +
                $"<Value><![CDATA[]]></Value></Property><Property path='followUpDate'><Value><![CDATA[{GetDateString()}]]></Value></Property>" +
                $"<Property path='recalculateDockDate'><Value><![CDATA[1]]></Value></Property><Property path='unitPriceRequested'>" +
                $"<Value><![CDATA[{id.UnitPrice}]]></Value></Property><Property path='item'><Value><![CDATA[{id.ItemNumber}]]></Value></Property>" +
                $"<Property path='charge'><Value><![CDATA[]]></Value></Property><Property path='warehouse'><Value><![CDATA[{id.Warehouse}]]></Value></Property>" +
                $"<Property path='planner'><Value><![CDATA[{id.Planner}]]></Value></Property><Property path='orderQuantityRequested'>" +
                $"<Value><![CDATA[{id.OrderQuantityRequested}]]></Value></Property><Property path='apportionment'>" +
                $"<Value><![CDATA[]]></Value></Property><Property path='safetyLeadTime'><Value><![CDATA[0]]></Value></Property>" +
                $"<Property path='countryOfOrigin'><Value><![CDATA[]]></Value></Property><Property path='dockToStockLeadTime'>" +
                $"<Value><![CDATA[0]]></Value></Property><Property path='orderUm'><Value><![CDATA[{id.UnitOfMeasure}]]></Value></Property>" +
                $"<Property path='userFieldSwitchA'><Value><![CDATA[N]]></Value></Property><Property path='printQuoteDescriptionsOnPo'>" +
                $"<Value><![CDATA[0]]></Value></Property><Property path='itemTaxTransactionType'><Value><![CDATA[]]></Value></Property>" +
                $"<Property path='reference'><Value><![CDATA[]]></Value></Property><Property path='nature'><Value><![CDATA[2069]]></Value></Property>" +
                $"<Property path='unit'><Value><![CDATA[]]></Value></Property><Property path='department'><Value><![CDATA[]]></Value></Property>" +
                $"<Property path='accountNumber'><Value><![CDATA[]]></Value></Property><Property path='relatedPoItemCommentExtension.extendedDescription2'>" +
                $"<Value><![CDATA[{id.ItemDescription}]]></Value></Property><Property path='relatedPoItemCommentExtension.extendedDescription1'>" +
                $"<Value><![CDATA[{id.ItemDescription}]]></Value></Property><Property path='requisition'><Value><![CDATA[]]></Value></Property>" +
                $"<Property path='jobNumber'><Value><![CDATA[]]></Value></Property><Property path='dueToDockDate'><Value><![CDATA[{GetDateString()}]]></Value></Property>" +
                $"<Property path='receiptRequired'><Value><![CDATA[1]]></Value></Property><Property path='taxPercent'><Value><![CDATA[0]]></Value></Property>" +
                $"<Property path='orderRescheduleCode'><Value><![CDATA[0]]></Value></Property><Property path='chargeType'><Value><![CDATA[]]></Value></Property>" +
                $"<Property path='vendorLeadTime'><Value><![CDATA[{id.VendorLeadTime}]]></Value></Property><Property path='drawingNumber'><Value><![CDATA[]]></Value></Property>" +
                $"<Property path='recalculateStockDate'><Value><![CDATA[1]]></Value></Property><Property path='advisePrice'><Value><![CDATA[0]]></Value></Property>" +
                $"<Property path='dueToStockDate'><Value><![CDATA[{GetDateString()}]]></Value></Property><Property path='blanketItem'><Value><![CDATA[0]]></Value></Property>" +
                $"<Property path='itemTaxClass'><Value><![CDATA[]]></Value></Property><Property path='relatedPoItemExtension.replicationDestination'><Value><![CDATA[]]></Value></Property>" +
                $"<Property path='fixedBlanket'><Value><![CDATA[0]]></Value></Property></DomainEntity><MaintainText name='maintainText_Comments' relationshipName='relatedPoItemComments'>" +
                $"<TextAction valueType='manual'><Type><![CDATA[P]]></Type><Value><![CDATA[]]></Value></TextAction><TextAction valueType='manual'><Type><![CDATA[C]]></Type>" +
                $"<Value><![CDATA[]]></Value></TextAction><TextAction valueType='manual'><Type><![CDATA[T]]></Type><Value><![CDATA[]]></Value></TextAction></MaintainText></Create></Request>" +
                $"</System-Link>\r\n";
            HttpWebResponse response = SendWebRequest(req);
            if (response.StatusCode == HttpStatusCode.OK)
            {
                Stream responseStream = response.GetResponseStream();
                string responseStr = new StreamReader(responseStream).ReadToEnd();
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(responseStr);
                string json = JsonConvert.SerializeXmlNode(doc, Newtonsoft.Json.Formatting.Indented);
                JObject resp = JObject.Parse(json);
            }
        }

        public void GetOrders()
        {
            string req = $"<?xml version='1.0' encoding='UTF-8'?><!DOCTYPE System-Link SYSTEM 'SystemLinkRequest.dtd'><System-Link>" +
                $"<Login userId='{Program.xaLogin.Key}' password='{Program.xaLogin.Value}' maxIdle='900000'\r\n" +
                $"          properties='com.pjx.cas.domain.EnvironmentId={userSettings.Environment.DisplayText.Substring(0, 2)},\r\n" +
                $"               com.pjx.cas.domain.SystemName=INFXAPRD,\r\n" +
                $"               com.pjx.cas.user.LanguageId=en'/><Request sessionHandle='*current' workHandle='*new'\r\n" +
                $"          broker='EJB' maxIdle='1000'>" +
                $"<QueryList name='queryListPurchaseOrder_GeneralR7' domainClass='com.mapics.pm.PurchaseOrder' includeMetaData='true' maxReturned='10'>" +
                $"     <Pql><![CDATA[SELECT order,priority,vendor,vendorName,\r\n" +
                $"                          orderStatus,holdFromPrint,approvalStatus,\r\n" +
                $"                          releaseDate,confirmByDate,confirmedDate,\r\n" +
                $"                          revision,lastRevisionDate,buyer,relatedBuyer.buyerName,\r\n" +
                $"                          warehouse WHERE orderCreatedDate > \r\n" +
                $"                          {DateTime.Now.AddDays(-7).ToString("yyyyMMdd")} ORDER BY order DESC]]></Pql></QueryList></Request></System-Link>\r\n";
            HttpWebResponse response = SendWebRequest(req);
            orders.Clear();
            if (response.StatusCode == HttpStatusCode.OK)
            {
                Stream responseStream = response.GetResponseStream();
                string responseStr = new StreamReader(responseStream).ReadToEnd();
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(responseStr);
                string json = JsonConvert.SerializeXmlNode(doc, Newtonsoft.Json.Formatting.Indented);
                JObject resp = JObject.Parse(json);
                JArray objects = (JArray)JsonConvert.DeserializeObject(resp["System-Link"]["Response"]["QueryListResponse"]["DomainEntity"].ToString());
                foreach (JObject obj in objects)
                {
                    Order order = new Order();
                    JArray properties = (JArray)JsonConvert.DeserializeObject(obj["Property"].ToString());
                    foreach (JObject property in properties)
                    {
                        switch (property["@path"].ToString())
                        {
                            case "order":
                                order.OrderNumber = GetValue(property);
                                break;
                            case "priority":
                                order.Details.Priority = GetValue(property);
                                break;
                            case "vendor":
                                order.Details.Vendor = GetValue(property);
                                break;
                            case "vendorName":
                                order.Details.VendorName = GetValue(property);
                                break;
                            case "orderStatus":
                                order.Details.OrderStatus = GetValue(property);
                                break;
                            case "holdFromPrint":
                                order.Details.HoldFromPrint = GetValue(property);
                                break;
                            case "approvalStatus":
                                order.Details.ApprovalStatus = GetValue(property);
                                break;
                            case "releaseDate":
                                order.Details.ReleaseDate = GetValue(property);
                                break;
                            case "confirmByDate":
                                order.Details.ConfirmByDate = GetValue(property);
                                break;
                            case "confirmedDate":
                                order.Details.ConfirmedDate = GetValue(property);
                                break;
                            case "revision":
                                order.Details.Revision = GetValue(property);
                                break;
                            case "lastRevisionDate":
                                order.Details.LastRevisionDate = GetValue(property);
                                break;
                            case "buyer":
                                order.Details.Buyer = GetValue(property);
                                break;
                            case "relatedBuyer.buyerName":
                                order.Details.BuyerName = GetValue(property);
                                break;
                            case "warehouse":
                                order.Details.Warehouse = GetValue(property);
                                break;
                        }

                    }

                    orders.Add(order);
                }
                poPreview.Clear();

                poPreview.Columns.Add(new ColumnHeader { Width = 78, Text = "Purchase Order" });
                poPreview.Columns.Add(new ColumnHeader { Width = 78, Text = "Vendor" });
                poPreview.Columns.Add(new ColumnHeader { Width = 78, Text = "Buyer" });

                foreach (Order order in orders)
                {
                    var item = new ListViewItem { Text = order.OrderNumber };
                    item.SubItems.Add(new ListViewItem.ListViewSubItem { Text = order.Details.Vendor });
                    item.SubItems.Add(new ListViewItem.ListViewSubItem { Text = order.Details.BuyerName });
                    poPreview.Items.Add(item);
                }
            }
        }

        public HttpWebResponse SendWebRequest(string xmlReq)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create("http://100.64.191.40:36001/SystemLink/servlet/SystemLinkServlet");
            byte[] bytes;
            bytes = Encoding.ASCII.GetBytes(xmlReq);
            request.ContentType = "text/xml; encoding='utf-8'";
            request.ContentLength = bytes.Length;
            request.Method = "POST";
            Stream requestStream = request.GetRequestStream();
            requestStream.Write(bytes, 0, bytes.Length);
            requestStream.Close();
            HttpWebResponse response;
            response = (HttpWebResponse)request.GetResponse();
            return response;
        }

        public string GetValue(JObject property)
        {
            if (property["Value"] is JObject)
                return property["Value"]["#cdata-section"].ToString();
            else
                return property["Value"].ToString();
        }

        public class ItemDetails
        {
            public string ItemNumber { get; set; }
            public string UnitPrice { get; set; }
            public string Warehouse { get; set; }
            public string Planner { get; set; }
            public string OrderQuantityRequested { get; set; }
            public string UnitOfMeasure { get; set; }
            public string ItemDescription { get; set; }
            public string VendorLeadTime { get; set; }
        }

        public class Order
        {
            public string OrderNumber { get; set; }
            public OrderDetails Details = new OrderDetails();

            public class OrderDetails
            {
                public string Priority { get; set; }
                public string Vendor { get; set; }
                public string VendorName { get; set; }
                public string OrderStatus { get; set; }
                public string HoldFromPrint { get; set; }
                public string ApprovalStatus { get; set; }
                public string ReleaseDate { get; set; }
                public string ConfirmByDate { get; set; }
                public string ConfirmedDate { get; set; }
                public string Revision { get; set; }
                public string LastRevisionDate { get; set; }
                public string Buyer { get; set; }
                public string BuyerName { get; set; }
                public string Warehouse { get; set; }
            }
        }

        private void browseBtn_Click(object sender, EventArgs e)
        {
            excelPathDiag = new OpenFileDialog();
            excelPathDiag.Title = "Select the Excel document to load:";
            excelPathDiag.Filter = "Excel files (*.xlsx)|*.xlsx";
            DialogResult result = excelPathDiag.ShowDialog();

            if (result == DialogResult.OK)
            {
                excelFilePath = excelPathDiag.FileName;
                textBox1.Text = excelFilePath;
                LoadDataToPreview();
            }
            else
            {
                //handle no file selected
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Program.xaLogin.Key == "")
            {
                Configuration login = new Configuration();
                login.ShowDialog();
            }

            else
            {
                GetOrders();
            }
        }

        private void excelDataPreview_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (excelDataPreview.SelectedIndices.Count > 0)
            {
                delItem.Enabled = true;
                excelViewIndex = excelDataPreview.SelectedIndices[0];
            }
        }

        private void delItem_Click(object sender, EventArgs e)
        {
            excelDataPreview.Items[excelViewIndex].Remove();
            delItem.Enabled = false;
        }

        private void poPreview_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (poPreview.SelectedItems.Count > 0)
            {
                poViewIndex = poPreview.SelectedIndices[0];
                uploadBtn.Enabled = true;
            }
        }

        private void uploadBtn_Click(object sender, EventArgs e)
        {
            uploadBtn.Enabled = false;
            BeginItemsUpload();
        }

        private string GetDateString() => DateTime.Now.ToString("yyyyMMdd");

        private void button2_Click(object sender, EventArgs e)
        {
            Configuration configuration = new Configuration();
            configuration.ShowDialog();
        }

        private void refreshBtn_Click(object sender, EventArgs e)
        {
            GetOrders();
        }
    }
}
