# =====================================================================
# PowerForm Full Demo Provisioning Script (Idempotent)
# Company: M365 Power Products
# =====================================================================

$SiteUrl = "https://m365powerproducts.sharepoint.com/sites/powerformdemo"
$ClientId = "48139e29-4085-44b5-8ef7-f4d47bb7c57a"

# 1. Connection & Module Setup
Write-Host "Checking for PnP.PowerShell..." -ForegroundColor Gray
if (-not (Get-Module -ListAvailable PnP.PowerShell)) {
    Write-Host "Installing PnP.PowerShell module..." -ForegroundColor Yellow
    Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force
}

Write-Host "Connecting to SharePoint..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

# --- HELPER FUNCTIONS ---

function Ensure-PnPList {
    param([string]$Title, [string]$Url)
    $list = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue
    if ($null -eq $list) {
        Write-Host "Creating List: $Title..." -ForegroundColor Cyan
        return New-PnPList -Title $Title -Template GenericList -Url $Url
    }
    Write-Host "List '$Title' already exists. Moving next..." -ForegroundColor Gray
    return $list
}

function Ensure-PnPField {
    param($List, [string]$DisplayName, [string]$InternalName, [string]$Type, $Params = @{}, $ExtraValues = @{})
    $field = Get-PnPField -List $List -Identity $InternalName -ErrorAction SilentlyContinue
    if ($null -eq $field) {
        Write-Host "  Adding Field: $DisplayName..." -ForegroundColor White
        $field = Add-PnPField -List $List -DisplayName $DisplayName -InternalName $InternalName -Type $Type -AddToDefaultView @Params
    }
    if ($ExtraValues.Count -gt 0) {
        Set-PnPField -List $List -Identity $InternalName -Values $ExtraValues
    }
    return $field
}

function Ensure-PnPListItem {
    param([string]$List, [hashtable]$Values)
    $existing = Get-PnPListItem -List $List -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($Values['Title'])</Value></Eq></Where></Query></View>"
    if ($existing.Count -eq 0) {
        Write-Host "  Adding Item: $($Values['Title']) to $List..." -ForegroundColor DarkGray
        Add-PnPListItem -List $List -Values $Values
    }
}

# --- PROVISIONING START ---

# 1. Vendors List
$vendorList = Ensure-PnPList -Title "Vendors" -Url "Lists/Vendors"
Ensure-PnPField -List "Vendors" -DisplayName "VendorEmail" -InternalName "VendorEmail" -Type Text
Ensure-PnPField -List "Vendors" -DisplayName "VendorRating" -InternalName "VendorRating" -Type Number
Ensure-PnPField -List "Vendors" -DisplayName "Active" -InternalName "Active" -Type Boolean # Added Active Column

Ensure-PnPListItem -List "Vendors" -Values @{"Title"="Microsoft"; "VendorEmail"="sales@microsoft.com"; "VendorRating"=5; "Active"=$true}
Ensure-PnPListItem -List "Vendors" -Values @{"Title"="Dell Technologies"; "VendorEmail"="b2b@dell.com"; "VendorRating"=4; "Active"=$true}
Ensure-PnPListItem -List "Vendors" -Values @{"Title"="Adobe"; "VendorEmail"="licensing@adobe.com"; "VendorRating"=4; "Active"=$true}

# 2. Categories & SubCategories
$catList = Ensure-PnPList -Title "Categories" -Url "Lists/Categories"
Ensure-PnPField -List "Categories" -DisplayName "Active" -InternalName "Active" -Type Boolean # Added Active Column
Ensure-PnPListItem -List "Categories" -Values @{"Title"="IT Hardware"; "Active"=$true}
Ensure-PnPListItem -List "Categories" -Values @{"Title"="Software Licensing"; "Active"=$true}
Ensure-PnPListItem -List "Categories" -Values @{"Title"="Facilities"; "Active"=$true}

$subList = Ensure-PnPList -Title "SubCategories" -Url "Lists/SubCategories"
Ensure-PnPField -List "SubCategories" -DisplayName "Active" -InternalName "Active" -Type Boolean # Added Active Column

# SubCategory Lookup to Categories (Fixed for String conversion)
if ($null -eq (Get-PnPField -List "SubCategories" -Identity "ParentCategory" -ErrorAction SilentlyContinue)) {
    Add-PnPField -List "SubCategories" -DisplayName "ParentCategory" -InternalName "ParentCategory" -Type Lookup -AddToDefaultView
}
Set-PnPField -List "SubCategories" -Identity "ParentCategory" -Values @{LookupList=$catList.Id.ToString(); LookupField="Title"}

# Link SubCategory Items
$itCat = (Get-PnPListItem -List "Categories" | Where-Object { $_["Title"] -eq "IT Hardware" }).Id
$swCat = (Get-PnPListItem -List "Categories" | Where-Object { $_["Title"] -eq "Software Licensing" }).Id

if ($null -ne $itCat) {
    Ensure-PnPListItem -List "SubCategories" -Values @{"Title"="Laptops"; "ParentCategory"=$itCat; "Active"=$true}
    Ensure-PnPListItem -List "SubCategories" -Values @{"Title"="Monitors"; "ParentCategory"=$itCat; "Active"=$true}
}
if ($null -ne $swCat) {
    Ensure-PnPListItem -List "SubCategories" -Values @{"Title"="Office 365"; "ParentCategory"=$swCat; "Active"=$true}
    Ensure-PnPListItem -List "SubCategories" -Values @{"Title"="Creative Cloud"; "ParentCategory"=$swCat; "Active"=$true}
}

# 3. Main List: Purchase Requests
Write-Host "Provisioning Main Purchase Requests List..." -ForegroundColor Cyan
$mainList = Ensure-PnPList -Title "Purchase Requests" -Url "Lists/PurchaseRequests"

# Autocomplete & Mapping Fields
Ensure-PnPField -List "Purchase Requests" -DisplayName "VendorName" -InternalName "VendorName" -Type Text -ExtraValues @{Description="PowerForm: Autocomplete."}
Ensure-PnPField -List "Purchase Requests" -DisplayName "VendorEmail" -InternalName "VendorEmail" -Type Text -ExtraValues @{Description="PowerForm: Column Mapping."}
Ensure-PnPField -List "Purchase Requests" -DisplayName "VendorRating" -InternalName "VendorRating" -Type Number -ExtraValues @{Description="PowerForm: Column Mapping."}

# Category Lookup (Fixed String ID)
if ($null -eq (Get-PnPField -List "Purchase Requests" -Identity "Category" -ErrorAction SilentlyContinue)) {
    Add-PnPField -List "Purchase Requests" -DisplayName "Category" -InternalName "Category" -Type Lookup -AddToDefaultView
}
Set-PnPField -List "Purchase Requests" -Identity "Category" -Values @{LookupList=$catList.Id.ToString(); LookupField="Title"; Description="PowerForm: Single Lookup."}

# SubCategory Lookup (Fixed String ID)
if ($null -eq (Get-PnPField -List "Purchase Requests" -Identity "SubCategory" -ErrorAction SilentlyContinue)) {
    Add-PnPField -List "Purchase Requests" -DisplayName "SubCategory" -InternalName "SubCategory" -Type Lookup -AddToDefaultView
}
Set-PnPField -List "Purchase Requests" -Identity "SubCategory" -Values @{LookupList=$subList.Id.ToString(); LookupField="Title"; Description="PowerForm: Cascade Lookup."}

# Choice & Standard Fields
Ensure-PnPField -List "Purchase Requests" -DisplayName "Status" -InternalName "Status" -Type Choice -Params @{Choices="Draft","Pending Approval","Approved","Rejected"}
Ensure-PnPField -List "Purchase Requests" -DisplayName "DeliveryLocations" -InternalName "DeliveryLocations" -Type MultiChoice -Params @{Choices="New York","London","Dubai","Tokyo"}
Ensure-PnPField -List "Purchase Requests" -DisplayName "TotalAmount" -InternalName "TotalAmount" -Type Currency
Ensure-PnPField -List "Purchase Requests" -DisplayName "IsUrgent" -InternalName "IsUrgent" -Type Boolean
Ensure-PnPField -List "Purchase Requests" -DisplayName "DateRequired" -InternalName "DateRequired" -Type DateTime

# FIX: Multi-Person Picker (Watchers)
Ensure-PnPField -List "Purchase Requests" -DisplayName "PrimaryContact" -InternalName "PrimaryContact" -Type User
Ensure-PnPField -List "Purchase Requests" -DisplayName "Watchers" -InternalName "Watchers" -Type User -ExtraValues @{AllowMultipleValues=$true; Description="PowerForm: Multi-Person Picker."}

# Links & Complex Types
Ensure-PnPField -List "Purchase Requests" -DisplayName "ReferenceLink" -InternalName "ReferenceLink" -Type URL
Ensure-PnPField -List "Purchase Requests" -DisplayName "Justification" -InternalName "Justification" -Type Note -ExtraValues @{RichText=$true; Description="PowerForm: Custom RTE."}
Ensure-PnPField -List "Purchase Requests" -DisplayName "Milestones" -InternalName "Milestones" -Type Note -ExtraValues @{RichText=$false; Description="PowerForm: Repeater Grid (JSON Data)."}

# 4. Child List: PO Line Items
$childList = Ensure-PnPList -Title "PO Line Items" -Url "Lists/POLineItems"
Ensure-PnPField -List "PO Line Items" -DisplayName "Quantity" -InternalName "Quantity" -Type Number
Ensure-PnPField -List "PO Line Items" -DisplayName "UnitPrice" -InternalName "UnitPrice" -Type Currency
Ensure-PnPField -List "PO Line Items" -DisplayName "PurchaseRequestId" -InternalName "PurchaseRequestId" -Type Number -ExtraValues @{Description="Foreign Key to Parent."}

Write-Host "Provisioning Complete! Ready for PowerForm Demo." -ForegroundColor Green