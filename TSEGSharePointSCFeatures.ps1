$SiteURL = "https://echoent.sharepoint.com/sites/Testprojectdelivery"
$SiteCollection = get-SPSite $SiteURL
Enable-PFeature "PublishingSite" -Url $siteCollection.Url -force
Enable-SPFeature "PublishingWeb" -Url $siteCollection.Url -force