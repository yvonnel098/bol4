#latest version of Bill of Lading development
#corresponding to original version 5
#
#This program imports customer address from server into Excel templates
#
#by Yvonne Lu
#
#changes made for this version:
#correct city, state, zip being blank for foreign address problem
#
#use File::Temp to safely get a temp file name
#
#copy BOL template and insert addresses then save to a temp file.  Open temp file in excel to avoid Excel prompt the
#user to save changed file at the first invokation. This is a shortcoming in previous version
#
#

use Tkx;
use Win32::ODBC;
use warnings;
use String::Util 'trim';
use Win32::OLE;
$Win32::OLE::Warn = 3; # Die on errors in Excel
use File::Temp qw(tempdir);
use autodie;

#address to pass to excel
$line1='';
$line2='';
$line3='';
$linePO = '';
#template parent directory - hard coded for now
$parent = "c:\\BOL\\";

#temporary directory that will delete
my $tmpdir = tempdir( CLEANUP => 1 );

#connect to db

my $dsn = "Driver={Microsoft Visual FoxPro Driver};SourceType=DBF;SourceDB=c:\\ShipInfo\\;Exclusive=NO;collate=Machine;NULL=NO;DELETED=NO;BACKGROUNDFETCH=NO;";
my $db= new Win32::ODBC($dsn) || die("cannot connect to database");
$db->Sql("SELECT sono, company FROM soaddr01 order by sono");
   
my $listnames = '';
my $query_count=0;
my $sononum;
while ($db->FetchRow()) {
    (my $listsono, my $company) = $db->Data("sono", "company");
    if ($query_count==0) {
	#$sononum = $listsono;
	$query_count++;
    }
    
    $listnames = $listnames . ' {' . $listsono.' '.$company . '}';
    
}

=for comment
my $test=395661;

$db->Sql("SELECT address1 FROM soaddr01 where sono like '%".$test."%'");
$db->FetchRow();
(my $addr1) = $db->Data("address1");
Tkx::tk___messageBox(-message => $addr1);
=cut



#get list of templates

opendir (DIR, $parent) or die $!;
my @templates 
        = grep { 
            m/\.xls$/             # Begins with a period
	    && -f "$parent/$_"   # and is a file
	} readdir(DIR);
        


my $temp_count = @templates;
closedir(DIR);

my $mw = Tkx::widget->new(".");
$mw->g_wm_title("Bill of Lading");
$mw->g_wm_iconbitmap('.\favicon.ico');
#$mw->protocol('WM_DELETE_WINDOW', [\&menubar_command, 'quit']);
$mw->g_wm_protocol( WM_DELETE_WINDOW => \&myexit);

#sono entry and listbox
my $sonolbl = $mw->new_ttk__label(-text => "SONO:");
$sonolbl->g_grid(-column => 0, -row => 0, -padx=>3);
my $sono = $mw->new_ttk__entry(-width => 30, -textvariable => \$sononum);
$sono->g_grid(-column=>1, -row=>0, -padx=>5);

($lb = $mw->new_tk__listbox(-listvariable => \$listnames, -height => 30))->g_grid(-column => 0, -row => 1, -columnspan => 2, -pady => 2, -sticky => "nwes");
#($lb = $mw->new_tk__listbox(-height => 5))->g_grid(-column => 0, -row => 0, -sticky => "nwes");
($s = $mw->new_ttk__scrollbar(-command => [$lb, "yview"], 
        -orient => "vertical"))->g_grid(-column =>2, -row => 1, -sticky => "ns");
$lb->configure(-yscrollcommand => [$s, "set"]);

#specific address and po
my $textbox = $mw->new_tk__text(-width => 40, -height => 20);
$textbox->g_grid(-column =>3, -row=>1, -padx => 10, -pady=>2);

#BOL template
my $lbl2 = $mw->new_ttk__label(-text => "BOL Templates:");
my $cnames = ''; foreach $i (@templates) {$cnames = $cnames . ' {' . $i . '}';};
my $lbox = $mw->new_tk__listbox(-listvariable => \$cnames, -width => 40, -height => $temp_count);
$lbl2->g_grid(-column => 4, -row => 0, -padx => 10, -pady => 5);
$lbox->g_grid(-column => 4, -row => 1, -padx => 10, -pady => 2, -sticky => "nswe");

#checkbox to includ PO number
my $po_inc=0;
my $po_but = $mw->new_ttk__checkbutton(-text => "Include PO#", -variable => \$po_inc, -onvalue => 1);
$po_but->g_grid(-column => 4, -row=>2, -pady=>3);

#button to start Excel
my $excel_but = $mw->new_ttk__button(-text => "Start Excel", -command => sub{invoke_excel();});
$excel_but->g_grid(-column => 4, -row=>3, -pady=>3);


#$sono->g_bind("<Return>",     get_address($db, $sononum));
$sono->g_bind("<Return>", sub {get_address($db, $sononum, $textbox, $sono);});
$lb->g_bind("<<ListboxSelect>>", sub {get_address_from_listbox($db, $textbox, $sono, $lb);});

Tkx::focus($sono);
Tkx::MainLoop();



sub get_address_from_listbox
{
    my $db = $_[0]; #database connection
    my $textbox = $_[1]; #textbox handle
    my $sono = $_[2]; #sono entry box handle
    my $lb = $_[3]; #list box handle
    
    #parse out sono number
    my @idx = $lb->curselection;
    my $seltext = $lb->get($idx[0]);
    my ($sononum, $company) = split /\s* \s*/, trim($seltext), 2;
    #Tkx::tk___messageBox(-message => $sononum.':::'.$company);
    get_address($db, $sononum, $textbox, $sono);
    
}

sub get_address
{
    my $db = $_[0]; #database connection
    my $sononum = trim($_[1]); #sono
    my $textbox = $_[2]; #textbox handle
    my $sono = $_[3];
    
    #Tkx::tk___messageBox(-message => "In get_address now ".$sononum);
    if ($sononum>0) {
	
    
	$db->Sql("SELECT company, address1, address2, address3, city, state, zip, cpo FROM soaddr01 where sono like '%".$sononum."'");
	#$db->Sql("SELECT company, address1, address2, address3, city, state, zip, cpo FROM soaddr01 where sono = '".$sononum."'");
	#$db->FetchRow() || die qq(Fetch error: ), $db->Error(), qq(\n);
	my $str="";
	if ($db->FetchRow()) {
	    (my $company, my $addr1, my $addr2, my $addr3, my $city, my $state, my $zip, my $cpo) =
	    $db->Data("company", "address1", "address2", "address3", "city", "state", "zip", "cpo");
	    #my $name=$company.chr(012).chr(015)."PO Number:  ".$cpo;
	    my $name=$company;
	    my $citystate = $city." ".$state." ".$zip;
	    
	    my $addr .=  trim($addr1) eq "" ? "" : $addr1.chr(012).chr(015);
	    $addr .=  trim($addr2) eq "" ? "" : $addr2.chr(012).chr(015);
	    $addr .= trim($addr3) eq "" ? "" : $addr3;
	    
	    $line1 = $name;
	    $line2 = $addr;
	    $line3 = $citystate;
	    $linePO = "PO Number:  ".$cpo;
	    
	    $str = $name.chr(012).chr(015).$linePO.chr(012).chr(015).$addr.$citystate;
	    
	    $textbox->delete("1.0", "end");
	    $textbox->insert("1.0", $str);
	    
	}
	else {
	    Tkx::tk___messageBox(-message => "No Address Found for SONO: ".$sononum);
	    $textbox->delete("1.0", "end");
	}
	
	#foreach $fd ($db->FieldNames())
	#{
	#    $str .= $db->Data($fd).chr(012).chr(015) unless (trim($db->Data($fd) eq''));
	    #print qq($fd: "), $db->Data($fd), qq("\n)};
	    
	#}
	#if ($count <=0) {
	#    Tkx::tk___messageBox(-message => "No Address Found for SONO: ".$sononum);
	#}
	#Tkx::tk___messageBox(-message => "count=".$count);
	
	#$textbox->insert("end", $addr1);
	#$textbox->insert("end", $addr2);
	#Tkx::tk___messageBox(-message => $str);
	
	
    }
    $sono->delete(0, "end");
    return;
    
    
}
    

sub invoke_excel
{
    
    #get temp file
    my $fh = File::Temp->new(DIR => $tmpdir);
    my $filename = $fh->filename.".xls"; #must have excel extension or else it won't work
    #my $filename="MyTmp1.xls";

   
    
    #note line3 (city, state, zip) may be blank for foreign addresses such as Canada
    if ((trim($line1) eq "") || (trim($line2)eq ""))
    {
	Tkx::tk___messageBox(-message => "Please select address first");
    }else
    {
	#get excel template
	my @idx = $lbox->curselection;
	if ($idx[0] eq "") {
	    Tkx::tk___messageBox(-message => "Please select a BOL template");
	} else {
	
	
	    my $template = $lbox->get($idx[0]);
	    #Tkx::tk___messageBox(-message => "template=".$template);
	    
	    # Start Excel and make it visible
	    my $Excel = Win32::OLE->GetActiveObject('Excel.Application')
		    || Win32::OLE->new('Excel.Application', 'QuitApp');
		    
	    #make it not visible
	    $Excel->{Visible} = 0;
	    
	    
	    my $Book = $Excel->Workbooks->Open($parent.$template);
    
    
	    my $Sheet = $Book->Worksheets("Sheet1");
	    $Sheet->Activate();
	    
	    #insert address
	    my $tmpline=$line1;
	    if ($po_inc) {
		$tmpline .= chr(012).chr(015).$linePO;
	    }
	    
	    $Sheet->Range("c9")->{Value} = $tmpline;
	    $Sheet->Range("c12")->{Value} = $line2;
	    $Sheet->Range("c14")->{Value} = $line3;
	    
	    #save to tmp file
	    $Book->SaveAs($filename);
	    $Book->Close();
	    
	    
	    #make it visible
	    $Excel->{Visible} = 1;
	    
	    #open temp file
	    my $tmpBook = $Excel->Workbooks->Open($filename);


	}
		

    }
    return;
    
    
}

sub QuitApp
{
    #Tkx::tk___messageBox(-message => "In QuitApp NOW");
    my ($ComObject)= @_;
    print "Qitting ".$ComObject->{Name}."\n";
    $ComObject->Quit();
}

sub myexit
{
    #Tkx::tk___messageBox(-message => "In My Exit Now");
    #see if excel has been closed
    my $Excel = Win32::OLE->GetActiveObject('Excel.Application');
    QuitApp($Excel);
    #close ODBC connection
    $db->Close();
    exit 1;
    
    
}