$Users = Get-ADUser -Properties *
foreach ($User in $users)
{
    Switch ($User.company)
    {
        "Aston Carter"              {$Company = "Aston Carter"}
        "AstonCarter"               {$Company = "Aston Carter"}
        "Teksystems"                {$Company = "TEKsystems" }
        "Tek systems"                {$Company = "TEKsystems" }
        "Aerotek"                   {$Company = "Aerotek"}
        "Allegis Global Solutions"  {$Company = "Allegis GLobal Solutions"}
        "AGS"                       {$Company = "Allegis GLobal Solutions"}
        "Allegis Partners"          {$Company = "Allegis Partners"}
        "AllegisPartners"           {$Company = "Allegis Partners"}
        "Allegis Group"             {$Company = "Allegis Group"}
        "AllegisGroup"              {$Company = "Allegis Group"}
    }
    set-aduser $user -company $Company
}
