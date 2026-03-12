# =====================================================================
# PowerForm Demo Provisioning Script
# =====================================================================
$SiteUrl = "https://YOUR_TENANT.sharepoint.com/sites/YOUR_SITE"

# Connect to SharePoint
Connect-PnPOnline -Url $SiteUrl -Interactive

Write-Host "Creating Supporting Lists..." -ForegroundColor Cyan

# 1. Vendors List (For Autocomplete & Mapping)
New-PnPList -Title "Vendors" -Template GenericList -Url "Lists/Vendors"
Add-PnPField -List "Vendors" -DisplayName "VendorEmail" -InternalName "VendorEmail" -Type Text -AddToDefaultView
Add-PnPField -List "Vendors" -DisplayName "VendorRating" -InternalName "VendorRating" -Type Number -AddToDefaultView

Add-PnPListItem -List "Vendors" -Values @{"Title"="Microsoft"; "VendorEmail"="sales@microsoft.com"; "VendorRating"=5}
Add-PnPListItem -List "Vendors" -Values @{"Title"="Dell Technologies"; "VendorEmail"="b2b@dell.com"; "VendorRating"=4}
Add-PnPListItem -List "Vendors" -Values @{"Title"="Adobe"; "VendorEmail"="licensing@adobe.com"; "VendorRating"=4}

# 2. Categories List (Parent Lookup)
$catList = New-PnPList -Title "Categories" -Template GenericList -Url "Lists/Categories"
Add-PnPListItem -List "Categories" -Values @{"Title"="IT Hardware"}
Add-PnPListItem -List "Categories" -Values @{"Title"="Software Licensing"}
Add-PnPListItem -List "Categories" -Values @{"Title"="Facilities"}

# 3. Sub-Categories List (Child Cascade Lookup)
$subList = New-PnPList -Title "SubCategories" -Template GenericList -Url "Lists/SubCategories"
Add-PnPField -List "SubCategories" -DisplayName "ParentCategory" -InternalName "ParentCategory" -Type Lookup -LookupListId $catList.Id -LookupField "Title" -AddToDefaultView

# Get Category IDs to link SubCategories
$itCat = (Get-PnPListItem -List "Categories" | Where-Object { $_["Title"] -eq "IT Hardware" }).Id
$swCat = (Get-PnPListItem -List "Categories" | Where-Object { $_["Title"] -eq "Software Licensing" }).Id

Add-PnPListItem -List "SubCategories" -Values @{"Title"="Laptops"; "ParentCategory"=$itCat}
Add-PnPListItem -List "SubCategories" -Values @{"Title"="Monitors"; "ParentCategory"=$itCat}
Add-PnPListItem -List "SubCategories" -Values @{"Title"="Office 365"; "ParentCategory"=$swCat}
Add-PnPListItem -List "SubCategories" -Values @{"Title"="Creative Cloud"; "ParentCategory"=$swCat}

Write-Host "Creating Main Purchase Requests List..." -ForegroundColor Cyan

# 4. Main List: Purchase Requests
$mainList = New-PnPList -Title "Purchase Requests" -Template GenericList -Url "Lists/PurchaseRequests"

# Add Fields
Add-PnPField -List "Purchase Requests" -DisplayName "VendorName" -InternalName "VendorName" -Type Text -AddToDefaultView
Add-PnPField -List "Purchase Requests" -DisplayName "VendorEmail" -InternalName "VendorEmail" -Type Text -AddToDefaultView
Add-PnPField -List "Purchase Requests" -DisplayName "VendorRating" -InternalName "VendorRating" -Type Number -AddToDefaultView
Add-PnPField -List "Purchase Requests" -DisplayName "Category" -InternalName "Category" -Type Lookup -LookupListId $catList.Id -LookupField "Title" -AddToDefaultView
Add-PnPField -List "Purchase Requests" -DisplayName "SubCategory" -InternalName "SubCategory" -Type Lookup -LookupListId $subList.Id -LookupField "Title" -AddToDefaultView
Add-PnPField -List "Purchase Requests" -DisplayName "Status" -InternalName "Status" -Type Choice -Choices "Draft","Pending Approval","Approved","Rejected" -AddToDefaultView
Add-PnPField -List "Purchase Requests" -DisplayName "DeliveryLocations" -InternalName "DeliveryLocations" -Type MultiChoice -Choices "New York","London","Dubai","Tokyo" -AddToDefaultView
Add-PnPField -List "Purchase Requests" -DisplayName "TotalAmount" -InternalName "TotalAmount" -Type Currency -AddToDefaultView
Add-PnPField -List "Purchase Requests" -DisplayName "IsUrgent" -InternalName "IsUrgent" -Type Boolean -AddToDefaultView
Add-PnPField -List "Purchase Requests" -DisplayName "DateRequired" -InternalName "DateRequired" -Type DateTime -AddToDefaultView
Add-PnPField -List "Purchase Requests" -DisplayName "PrimaryContact" -InternalName "PrimaryContact" -Type User -AddToDefaultView
Add-PnPField -List "Purchase Requests" -DisplayName "Watchers" -InternalName "Watchers" -Type UserMulti -AddToDefaultView
Add-PnPField -List "Purchase Requests" -DisplayName "ReferenceLink" -InternalName "ReferenceLink" -Type URL -AddToDefaultView

# Add Rich Text Field
$justField = Add-PnPField -List "Purchase Requests" -DisplayName "Justification" -InternalName "Justification" -Type Note -AddToDefaultView
Set-PnPField -Identity $justField -Values @{RichText=$true}

# Add Plain Text Field (For the JSON Repeater Grid)
$milestoneField = Add-PnPField -List "Purchase Requests" -DisplayName "Milestones" -InternalName "Milestones" -Type Note -AddToDefaultView
Set-PnPField -Identity $milestoneField -Values @{RichText=$false}

# 5. Child List (For Line Items)
Write-Host "Creating Child List..." -ForegroundColor Cyan
$childList = New-PnPList -Title "PO Line Items" -Template GenericList -Url "Lists/POLineItems"
Add-PnPField -List "PO Line Items" -DisplayName "Quantity" -InternalName "Quantity" -Type Number -AddToDefaultView
Add-PnPField -List "PO Line Items" -DisplayName "UnitPrice" -InternalName "UnitPrice" -Type Currency -AddToDefaultView
Add-PnPField -List "PO Line Items" -DisplayName "PurchaseRequestId" -InternalName "PurchaseRequestId" -Type Number -AddToDefaultView # Foreign Key

Write-Host "✅ Provisioning Complete! You can now configure your web part." -ForegroundColor Green