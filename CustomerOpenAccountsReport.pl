##
## 
use Time::Piece;
use Math::Round;
use IO::Handle;
use Cwd;

my $timestart;
my $timeend;

 my @AoA;  # Full listing
 my @BoB;  # Score Sort
 my @CoC;  # Last Payment Date Sort
 my @DoD;  # Insurance Expiration Sort
 my @EoE;  # Days Late Sort
 my @FoF;  # Days Late + Last Payment Sort
 
 
 my $REPOSSESSED = "Reposessed";
 my $INPROCESSOFREPO = "In Process of Repo";
 my $ONHOLD = "Hold";
 my $INS_SCORE_WEIGHT = 5.5;
 my $LAST_PAYMENT_WEIGHT = 5;
 my $DAYS_LATE_WEIGHT = .25;
 my $TOP_PRIORITY_WEIGHT = 1000;
 my $INSEXP_SKIP = 500;
 
 use constant {
	ACCT_DAYSLATE_INDEX   => 0,
	ACCT_LASTPAYMENT_INDEX => 1,
	ACCT_LASTPAYMENT_DELTA_INDEX => 2,
	ACCT_NAME_INDEX  => 3,
	ACCT_REPO_INDEX  => 4,
	ACCT_INSEXPIRE_INDEX => 5,
	ACCT_INSEXPIRE_DELTA_INDEX => 6,
	ACCT_CAR_INDEX => 7,
	ACCT_SCORE_INDEX => 8,
	ACCT_SALE_DATE_DELTA_INDEX => 9,
	ACCT_LATE_COMBO_INDEX => 10,
	ACCT_PAYOFF_INDEX => 11,
	ACCT_HOME_PHONE_INDEX => 12,
	ACCT_CELL_PHONE_INDEX => 13,
	ACCT_WORK_PHONE_INDEX => 14,
	ACCT_TOTAL_DUE_INDEX => 15,
	ACCT_PAYMENTS_DUE_INDEX => 16,
	MAX_INDEX_VALUE => 17
};

# Get current date/time
$myToday = localtime;

my $myTodayFormat = localtime->strftime('%Y_%m_%d');
mkdir $myTodayFormat;

# CompleteOpenAccounts.csv
if (open(AM_INPUT_FILE,$ARGV[0]) == 0) {
   print "Error opening input AutoManager report file: ",$ARGV[0];
   exit -1;  
}

# InsuranceExpirationReport.csv
if (open(AM_INSURANCE_FILE,$ARGV[1]) == 0) {
   print "Error opening input insurance file: ",$ARGV[1];
   exit -1;  
}

my $ROOT_DIR = $ARGV[2];
my $CWD = getcwd;

my $overviewFilename = sprintf("%s\\OpenAcctOverview_%s.html",$myTodayFormat,$myTodayFormat);
my $filename = sprintf(">%s",$overviewFilename);
if (open(HTML_OUTPUT_FILE_OVERVIEW,$filename) == 0) {
   print "Error opening: %s",$filename,"\n";
   exit -1;  
}

my $scoreFilename = sprintf("%s\\OpenAcctScore_%s.html",$myTodayFormat,$myTodayFormat);
$filename = sprintf(">%s",$scoreFilename);
if (open(HTML_OUTPUT_FILE_SCORE,$filename) == 0) {
   print "Error opening: %s",$filename,"\n";
   exit -1;  
}

my $lateWeightedFilename = sprintf("%s\\OpenAcctComboLatePayment_%s.html",$myTodayFormat,$myTodayFormat);
$filename = sprintf(">%s",$lateWeightedFilename);
if (open(HTML_OUTPUT_FILE_COMBO_LATE,$filename) == 0) {
   print "Error opening: %s",$filename,"\n";
   exit -1;
} 

$filename = sprintf(">%s\\OpenAcctLatePayment_%s.html",$myTodayFormat,$myTodayFormat);
if (open(HTML_OUTPUT_FILE_LASTPAYMENT,$filename) == 0) {
   print "Error opening: %s",$filename,"\n";
   exit -1;  
}

$filename = sprintf(">%s\\OpenAcctInsurance_%s.html",$myTodayFormat,$myTodayFormat);
if (open(HTML_OUTPUT_FILE_INSURANCE,$filename) == 0) {
   print "Error opening: %s",$filename,"\n";
   exit -1;  
}

$filename = sprintf(">%s\\OpenAcctDaysLate_%s.html",$myTodayFormat,$myTodayFormat);
if (open(HTML_OUTPUT_FILE_DAYSLATE,$filename) == 0) {
   print "Error opening: %s",$filename,"\n";
   exit -1;  
}

my $totalNumOpenAccounts = 0;
my $numAcctsRepossessed = 0;
my $numAcctsOnHold=0;
my $myOpenAccountIndex = 0;
my $numAccountsSkipped = 0;
my $myFoundMatchingInsVin = 0;	
my $myMaxScoreIndex = 0;
my $myMaxDaysLate = 0;	
my $myMaxLastPaymentDate = 0;
my $myMaxDaysLateCombo = 0;
my $myMaxInsuranceExpiration = 0;
my $myMaxLateAndInsExpired = 0;
my $myEndofFile = 0;

my $totalNumOfOpenAccts = 0;
my $totalAmountFinanced = 0;
my $totalPayoffAmount = 0;
my $totalMontlyPaymentAmount = 0;

my $numAccts0DaysLate = 0;
my $numAccts1to9DaysLate = 0;
my $numAccts10to29DaysLate = 0;
my $numAccts30to44DaysLate = 0;
my $numAccts45to59DaysLate = 0;
my $numAccts60to89DaysLate = 0;
my $numAccts90to120DaysLate = 0;
my $numAccts120to180DaysLate = 0;
my $numAccts180to365DaysLate = 0;
my $numAccts365DaysLateOrGreater = 0;

my $numAccts0to30DaysSinceLastPayments = 0;
my $numAccts30to44DaysSinceLastPayment = 0;
my $numAccts45to59DaysSinceLastPayment = 0;
my $numAccts60to89DaysSinceLastPayment = 0;
my $numAccts90to120DaysSinceLastPayment = 0;
my $numAccts120to180DaysSinceLastPayment = 0;
my $numAccts180to365DaysSinceLastPayment = 0;
my $numAccts365DaysLateOrGreaterSinceLastPayment = 0;

# Write out to NewCompleteOpenAccounts.csv
$filename = sprintf(">%s\\NewCompleteAccountsOverview.csv",$myTodayFormat);
if (open(NEW_AM_INPUT_FILE,$filename) == 0) {
   print "Error opening input AutoManager report file: ",$filename,"\n";
   exit -1;  
}
while (<AM_INPUT_FILE>) 
{
	chomp;
	($autoyear,$vin,$totaldue,$dealdate,$homephone,$payoff,$monthlypayment,$automodel,$automake,$dayslate,$lastpaymentdate,$textOk,$lastname,$firstname,$repostatus,$cellphone,$amtfinanced,$adjustbalance,$lastpaymentdate,$workphone,$paymentsdue) = split(",");
	print NEW_AM_INPUT_FILE $autoyear,",",$vin,",",$totaldue,",",$dealdate,",",$homephone,",",$payoff,",",$monthlypayment,",",$automodel,",",$automake,",",$dayslate,",",$lastpaymentdate,",",$textOk,",",$lastname,",",$firstname,",",$repostatus,",",$cellphone,",",$amtfinanced,",",$adjustbalance,",",$lastpaymentdate,",",$workphone,",",$paymentsdue,"\n";
};
NEW_AM_INPUT_FILE->flush();
close(AM_INPUT_FILE); 
close(NEW_AM_INPUT_FILE);
print "NewCompleteAccountsOverview copy success!\n"; 

# Write out to NewInsuranceExpirationReport.csv
$filename = sprintf(">%s\\NewInsuranceExpirationReport.csv",$myTodayFormat);
if (open(NEW_AM_INSURANCE_FILE,$filename) == 0) {
   print "Error opening input insurance file: ",$filename,"\n";
   exit -1;  
}
while (<AM_INSURANCE_FILE>) 
{
	chomp;
	($saledate,$custname,$inscompany,$insbroker,$insstart,$insexpire,$policynum,$unused1,$unused2,$insvin,$unused4) = split(",");
	print NEW_AM_INSURANCE_FILE $saledate,",",$custname,",",$inscompany,",",$insbroker,",",$insstart,",",$insexpire,",",$policynum,",",$unused1,",",$unused2,",",$insvin,",",$unused4,"\n";
}
NEW_AM_INSURANCE_FILE->flush();
close(AM_INSURANCE_FILE); 
close(NEW_AM_INSURANCE_FILE);
print "NewInsuranceExpirationReport copy success!\n"; 

# Reopen for reading - NewCompleteOpenAccounts.csv
$filename = sprintf("<%s\\NewCompleteAccountsOverview.csv",$myTodayFormat);
if (open(NEW_AM_INPUT_FILE,$filename) == 0) {
   print "Error opening input AutoManager report file: ",$filename,"\n";
   exit -1;  
}

# Reopen for reading - NewInsuranceExpirationReport.csv
$filename = sprintf("<%s\\NewInsuranceExpirationReport.csv",$myTodayFormat);
if (open(NEW_AM_INSURANCE_FILE,$filename) == 0) {
   print "Error opening input insurance file: ",$filename,"\n";
   exit -1;  
}

while (<NEW_AM_INPUT_FILE>) 
{
	chomp;
	($autoyear,$vin,$totaldue,$dealdate,$homephone,$payoff,$monthlypayment,$automodel,$automake,$dayslate,$lastpaymentdate,$textOk,$lastname,$firstname,$repostatus,$cellphone,$amtfinanced,$adjustbalance,$lastpaymentdate,$workphone,$paymentsdue) = split(",");
	
	if (($repostatus eq $REPOSSESSED) || ($repostatus eq $INPROCESSOFREPO))
	{	
		$numAcctsRepossessed++;		
	}
	elsif ($repostatus eq $ONHOLD)
	{
		$numAcctsOnHold++;	
	}
	else
	{
		my $AoADaysLate = $dayslate;		

		if ($AoADaysLate >= 1 && $AoADaysLate < 10)
		{
			$numAccts1to9DaysLate++;
		}
		elsif ($AoADaysLate >= 10 && $AoADaysLate < 30)
		{
			$numAccts10to29DaysLate++;
		}
		elsif ($AoADaysLate >= 30 && $AoADaysLate < 45)
		{
			$numAccts30to44DaysLate++;
		}
		elsif ($AoADaysLate >= 45 && $AoADaysLate < 60)
		{
			$numAccts45to59DaysLate++;
		}
		elsif ($AoADaysLate >= 60 && $AoADaysLate < 90)
		{
			$numAccts60to89DaysLate++;
		}
		elsif ($AoADaysLate >= 90 && $AoADaysLate < 120)
		{
			$numAccts90to120DaysLate++;
		}
		elsif ($AoADaysLate >= 120 && $AoADaysLate < 180)
		{
			$numAccts120to180DaysLate++;
		}
		elsif ($AoADaysLate >= 180 && $AoADaysLate < 365) 
		{
			$numAccts180to365DaysLate++;
		}
		elsif ($AoADaysLate >= 365)
		{
			$numAccts365DaysLateOrGreater++;
		}
		else 
		{
			$numAccts0DaysLate++;
		}
		
		my $lpDate = Time::Piece->strptime($lastpaymentdate, "%m/%d/%yy");	
		my $lpDelta = ($myToday - $lpDate)/86400;
		
		if ($lpDelta >= 0 && $lpDelta < 30)
		{
			$numAccts0to30DaysSinceLastPayments++;
		}
		elsif ($lpDelta >= 30 && $lpDelta < 45)
		{
			$numAccts30to44DaysSinceLastPayment++;
		}
		elsif ($lpDelta >= 45 && $lpDelta < 60)
		{
			$numAccts45to59DaysSinceLastPayment++;
		}
		elsif ($lpDelta >= 60 && $lpDelta < 90)
		{
			$numAccts60to89DaysSinceLastPayment++;
		}
		elsif ($lpDelta >= 90 && $lpDelta < 120)
		{
			$numAccts90to120DaysSinceLastPayment++;
		}
		elsif ($lpDelta >= 120 && $lpDelta < 180)
		{
			$numAccts120to180DaysSinceLastPayment++;
		}
		elsif ($lpDelta >= 180 && $lpDelta < 365)
		{
			$numAccts180to365DaysSinceLastPayment++;
		}
		else
		{
			$numAccts365DaysLateOrGreaterSinceLastPayment++;
		}
	
		$totalNumOpenAccounts++;
	}
	
	$totalAmountFinanced += $amtfinanced;
	$totalPayoffAmount += ($payoff);
	$totalMontlyPaymentAmount += $monthlypayment;
	$totalAmountDue += $totaldue;
	
	my $saleDate = Time::Piece->strptime($dealdate, "%m/%d/%yy");

	$AoA[$myOpenAccountIndex][ACCT_DAYSLATE_INDEX]          = $dayslate; 
	$AoA[$myOpenAccountIndex][ACCT_CAR_INDEX]               = sprintf("%s %s %s",$autoyear,$automake,$automodel); 
	$AoA[$myOpenAccountIndex][ACCT_LASTPAYMENT_INDEX]       = $lastpaymentdate; 
	$AoA[$myOpenAccountIndex][ACCT_NAME_INDEX]              = sprintf("%s, %s",$lastname,$firstname);
    $AoA[$myOpenAccountIndex][ACCT_REPO_INDEX]              = $repostatus;
	$AoA[$myOpenAccountIndex][ACCT_SALE_DATE_DELTA_INDEX]   = sprintf("%d",($myToday - $saleDate)/86400);	
	$AoA[$myOpenAccountIndex][ACCT_PAYOFF_INDEX]            = $payoff;
	$AoA[$myOpenAccountIndex][ACCT_HOME_PHONE_INDEX]        = $homephone;
	$AoA[$myOpenAccountIndex][ACCT_CELL_PHONE_INDEX]        = $cellphone;
	$AoA[$myOpenAccountIndex][ACCT_WORK_PHONE_INDEX]        = $workphone;
	$AoA[$myOpenAccountIndex][ACCT_TOTAL_DUE_INDEX]         = $totaldue;
	$AoA[$myOpenAccountIndex][ACCT_PAYMENTS_DUE_INDEX]      = $paymentsdue;
	
	#Set Defaults
	$AoA[$myOpenAccountIndex][ACCT_INSEXPIRE_INDEX] = "ERROR";
    $AoA[$myOpenAccountIndex][ACCT_INSEXPIRE_DELTA_INDEX]   = 0;	
	$AoA[$myOpenAccountIndex][ACCT_LASTPAYMENT_DELTA_INDEX] = 0;
	$AoA[$myOpenAccountIndex][ACCT_SCORE_INDEX] = "ERROR";
	$AoA[$myOpenAccountIndex][ACCT_LATE_COMBO_INDEX] = "ERROR";
	
	$firstname =~ s/^ *//;
	$lastname =~ s/^ *//;
	$firstname = sprintf("%s",uc($firstname));
	$lastname  = sprintf("%s",uc($lastname));
	$myCustomerName = sprintf("%s %s",$lastname,$firstname);
	
	$myEndofFile = 0;
	seek(NEW_AM_INSURANCE_FILE,0,0);
	
	while (<NEW_AM_INSURANCE_FILE>) 
	{
		chomp;
		($saledate,$custname,$inscompany,$insbroker,$insstart,$insexpire,$policynum,$unused1,$unused2,$insvin,$unused4) = split(",");

		$custname =~ s/^ *//;
		$custname = sprintf("%s",uc($custname));
		
		if ( !length($custname) )
		{
			$myEndofFile = 1;
			last;
		}
		
		# Check for valid VIN len
		if ( length($insvin) ne 17 )
		{
			print "=> ERROR: Customer: [", $custname, "] INSURANCE VIN: [",$insvin,"] INVALID LENGTH!\n";			
		}
		
		my $insExpDate; 	
		my $insExpDelta;		
		
		$myFoundMatchingInsVin = 0;
			
		if ( $insvin eq $vin )
		{		
			$myFoundMatchingInsVin = 1;	
		
			if (length($insexpire) eq 0)
			{
				print "ERROR: NO INSURANCE EXPIRATION: ", $custname,"\n";
				$insExpDelta = 0;
			}
			else
			{
				$insExpDate = Time::Piece->strptime($insexpire, "%m/%d/%yy");	
				$insExpDelta = ($myToday - $insExpDate)/86400;		
			}
			
			my $lpDate = Time::Piece->strptime($lastpaymentdate, "%m/%d/%yy");	
			my $lpDelta = ($myToday - $lpDate)/86400;			
				
			if ($insExpDelta < 0 )
			{
				$insExpDelta = 0;
			}
			
			if ( $dayslate eq 0 )
			{
				$lpDelta = 0;
			}
			
			my $lpWeight = $LAST_PAYMENT_WEIGHT;
			my $insWeight = $INS_SCORE_WEIGHT;
			my $dlWeight = $DAYS_LATE_WEIGHT;
			
			my $saledate = Time::Piece->strptime($saledate, "%m/%d/%yy");
			my $saleDateDelta = ($myToday - $saledate)/86400;
			
			# New accounts that are already late, go on top of the list!
			if ((($saleDateDelta <= 90) && ($dayslate > 14)) )
			{
				printf $custname,"\n";
				$lpWeight = $TOP_PRIORITY_WEIGHT;
			}
			else 
			{
				# If last payment was made within 3 weeks, half the weight of last payment
				if ( $dayslate < 28 )
				{
					$lpWeight = $lpWeight/4;
				}
				
				if ( $dayslate < 21 )
				{
					$lpWeight = $lpWeight/3;
				}
				
				if ( $lpDelta < 14 )
				{
					$lpWeight = $lpWeight/2;
				}
			}
			
			# If last payment was made within 2 weeks, half the weight of the insurance expiration.
			if ( $insExpDelta <= 10 )
			{
				$insWeight = $insWeight/2;
			}			
					
			my $score = sprintf("%d",($dayslate*$dlWeight) + ($lpDelta*$lpWeight) + ($insExpDelta*$insWeight));
			my $lateComboScore;
			
			# If Insurance is greater than 10 days past due; then, add in an insurance weight to total.
			if ( $insExpDelta > 10 )
			{
				$lateComboScore = sprintf("%d",($dayslate*$dlWeight) + ($lpDelta*$lpWeight) + ($insExpDelta*$insWeight));
			}
			else
			{
				$lateComboScore = sprintf("%d",($dayslate*$dlWeight) + ($lpDelta*$lpWeight));
			}
			
			$AoA[$myOpenAccountIndex][ACCT_INSEXPIRE_INDEX] = $insexpire;
			$AoA[$myOpenAccountIndex][ACCT_INSEXPIRE_DELTA_INDEX]   = sprintf("%d",$insExpDelta);	
			$AoA[$myOpenAccountIndex][ACCT_LASTPAYMENT_DELTA_INDEX] = sprintf("%d",$lpDelta);
			$AoA[$myOpenAccountIndex][ACCT_SCORE_INDEX] = $score;
			$AoA[$myOpenAccountIndex][ACCT_LATE_COMBO_INDEX] = $lateComboScore;

			if ($dayslate > $myMaxDaysLate)
			{
				$myMaxDaysLate = $dayslate;
			}
			
			if ($lpDelta > $myMaxLastPaymentDate)
			{
				$myMaxLastPaymentDate = $lpDelta;
			}
			
			if ($insExpDelta > $myMaxInsuranceExpiration)
			{
				$myMaxInsuranceExpiration = $insExpDelta;
			}

			if ($score > $myMaxLateAndInsExpired)
			{
				$myMaxLateAndInsExpired = $score;
			}
			
			if ($lateComboScore > $myMaxDaysLateCombo)
			{
				$myMaxDaysLateCombo = $lateComboScore;
			}
		}			
	}
	
	
	if (($myFoundMatchingInsVin eq 0) && ($myEndofFile eq 0))
	{
		print "ERROR: UNABLE TO FIND INSURANCE FOR ACCT: ", $custname,"\n";
	}
	
	$myOpenAccountIndex++;			
}

# Adjust counter minus one since it was incremented 
# before while() EoF check.
$totalNumOpenAccounts--;

close(NEW_AM_INPUT_FILE); 
close(NEW_AM_INSURANCE_FILE);
print "File input into memory success!\n";

my $myScoreOutput     = HTML_OUTPUT_FILE_SCORE;
my $myLastPayOutput   = HTML_OUTPUT_FILE_LASTPAYMENT;
my $myInsOutput       = HTML_OUTPUT_FILE_INSURANCE;
my $myDaysLateOutput  = HTML_OUTPUT_FILE_DAYSLATE;
my $myLateComboOutput = HTML_OUTPUT_FILE_COMBO_LATE;
my $myOverviewOutput  = HTML_OUTPUT_FILE_OVERVIEW;

WriteHtmlTopPage($myScoreOutput,"Provident Open Accounts Report","SCORE");
WriteHtmlTopPage($myLastPayOutput,"Provident Open Accounts Report","LAST PAYMENT DATE");
WriteHtmlTopPage($myInsOutput,"Provident Open Accounts Report","INSURANCE EXPIRATION");
WriteHtmlTopPage($myDaysLateOutput,"Provident Open Accounts Report","DAYS LATE");
WriteHtmlTopPage($myLateComboOutput,"Provident Open Accounts Report","WEIGHTED-LATE");
WriteHtmlTopPage($myOverviewOutput,"Provident Open Accounts Report","OVERVIEW");

my $arrayIndex = 0;
$myMaxLastPaymentDate     = sprintf("%d",$myMaxLastPaymentDate);
$myMaxInsuranceExpiration = sprintf("%d",$myMaxInsuranceExpiration);
$myMaxLateAndInsExpired   = sprintf("%d",$myMaxLateAndInsExpired);

print "SIZE OF AoA (All Accounts): ",$#AoA,"\n";

########################
# BoB is sorted by SCORE
########################

for my $i (0 .. $#AoA) 
{
	for my $j (0 .. MAX_INDEX_VALUE-1)
	{
		$BoB[$i][$j] = $AoA[$i][$j];
	}	
}

my $changeMade = 1;
my $numChangesMade = 0;

while ($changeMade eq 1)
{
	$changeMade = 0;
	
	for my $i (0 .. ($#BoB - 1)) 
	{
		my $currentIndex = $i;
		my $nextIndex  = $i+1;
		
		my $aoaScore     = $BoB[$currentIndex][ACCT_SCORE_INDEX];		
		my $aoaScoreNext = $BoB[$nextIndex][ACCT_SCORE_INDEX];
		
		if ( $aoaScoreNext < $aoaScore )
		{
			my @ArrayCopy;  
			
			for my $j (0 .. MAX_INDEX_VALUE-1)
			{
				$ArrayCopy[$j] = $BoB[$i][$j];
			}
			
			for my $j (0 .. MAX_INDEX_VALUE-1)
			{
				$BoB[$i][$j] = $BoB[$nextIndex][$j];
			}
			
			for my $j (0 .. MAX_INDEX_VALUE-1)
			{
				$BoB[$nextIndex][$j] = $ArrayCopy[$j];
			}
			
			$changeMade = 1;
			$numChangesMade++;
		}
	}

}	

print "SIZE OF BoB (Sorted by Score Accounts): ",$#BoB,"\n";


###################################
#CoC is sorted by LAST PAYMENT DATE
###################################
$arrayIndex=0;

for (my $j = $myMaxLastPaymentDate,$arrayIndex=0; $j >=0; $j--)
{
    for my $i (0 .. $#AoA) 
	{		
		if ( ($AoA[$i][ACCT_LASTPAYMENT_DELTA_INDEX] eq $j) && ($AoA[$i][ACCT_INSEXPIRE_DELTA_INDEX] < $INSEXP_SKIP) )
		{
			$CoC[$arrayIndex][ACCT_DAYSLATE_INDEX]          = $AoA[$i][ACCT_DAYSLATE_INDEX];
			$CoC[$arrayIndex][ACCT_CAR_INDEX]               = uc($AoA[$i][ACCT_CAR_INDEX]);
			$CoC[$arrayIndex][ACCT_LASTPAYMENT_INDEX]       = $AoA[$i][ACCT_LASTPAYMENT_INDEX];
			$CoC[$arrayIndex][ACCT_NAME_INDEX]              = $AoA[$i][ACCT_NAME_INDEX]; 
			$CoC[$arrayIndex][ACCT_REPO_INDEX]              = $AoA[$i][ACCT_REPO_INDEX];
			$CoC[$arrayIndex][ACCT_INSEXPIRE_INDEX]         = $AoA[$i][ACCT_INSEXPIRE_INDEX];
            $CoC[$arrayIndex][ACCT_INSEXPIRE_DELTA_INDEX]   = $AoA[$i][ACCT_INSEXPIRE_DELTA_INDEX];	
			$CoC[$arrayIndex][ACCT_LASTPAYMENT_DELTA_INDEX] = $AoA[$i][ACCT_LASTPAYMENT_DELTA_INDEX];
			$CoC[$arrayIndex][ACCT_SCORE_INDEX]             = $AoA[$i][ACCT_SCORE_INDEX];
			$CoC[$arrayIndex][ACCT_SALE_DATE_DELTA_INDEX]   = $AoA[$i][ACCT_SALE_DATE_DELTA_INDEX];
			$CoC[$arrayIndex][ACCT_PAYOFF_INDEX]            = $AoA[$i][ACCT_PAYOFF_INDEX];	
			$arrayIndex++;
			next;
		}		
   } 
}

print "SIZE OF CoC (Sorted by Last Payment Date): ",$#CoC,"\n";

############################################
#DoD is sorted by INSURANCE EXPIRATION DELTA
############################################
$arrayIndex=0;

for (my $j = $myMaxInsuranceExpiration,$arrayIndex=0; $j >=0; $j--)
{
    for my $i (0 .. $#AoA) 
	{
		if ( $AoA[$i][ACCT_INSEXPIRE_DELTA_INDEX] eq $j )
		{
			$DoD[$arrayIndex][ACCT_DAYSLATE_INDEX]          = $AoA[$i][ACCT_DAYSLATE_INDEX];
			$DoD[$arrayIndex][ACCT_CAR_INDEX]               = uc($AoA[$i][ACCT_CAR_INDEX]);
			$DoD[$arrayIndex][ACCT_LASTPAYMENT_INDEX]       = $AoA[$i][ACCT_LASTPAYMENT_INDEX];
			$DoD[$arrayIndex][ACCT_NAME_INDEX]              = $AoA[$i][ACCT_NAME_INDEX]; 
			$DoD[$arrayIndex][ACCT_REPO_INDEX]              = $AoA[$i][ACCT_REPO_INDEX];
			$DoD[$arrayIndex][ACCT_INSEXPIRE_INDEX]         = $AoA[$i][ACCT_INSEXPIRE_INDEX];
            $DoD[$arrayIndex][ACCT_INSEXPIRE_DELTA_INDEX]   = $AoA[$i][ACCT_INSEXPIRE_DELTA_INDEX];	
			$DoD[$arrayIndex][ACCT_LASTPAYMENT_DELTA_INDEX] = $AoA[$i][ACCT_LASTPAYMENT_DELTA_INDEX];
			$DoD[$arrayIndex][ACCT_SCORE_INDEX]             = $AoA[$i][ACCT_SCORE_INDEX];
			$DoD[$arrayIndex][ACCT_SALE_DATE_DELTA_INDEX]   = $AoA[$i][ACCT_SALE_DATE_DELTA_INDEX];
			$DoD[$arrayIndex][ACCT_PAYOFF_INDEX]            = $AoA[$i][ACCT_PAYOFF_INDEX];	
			$arrayIndex++;
			next;
		}
	} 
}

print "SIZE OF DoD (Sorted by Insurance Expiration): ",$#DoD,"\n";


############################################
#EoE is sorted by PAYMENT DAYS LATE
############################################
$arrayIndex=0;

for (my $j = $myMaxDaysLate,$arrayIndex=0; $j >=0; $j--)
{
    for my $i (0 .. $#AoA) 
	{
		if ( ($AoA[$i][ACCT_DAYSLATE_INDEX] eq $j) && ($AoA[$i][ACCT_INSEXPIRE_DELTA_INDEX] < $INSEXP_SKIP) )
		{
			$EoE[$arrayIndex][ACCT_DAYSLATE_INDEX]          = $AoA[$i][ACCT_DAYSLATE_INDEX];
			$EoE[$arrayIndex][ACCT_CAR_INDEX]               = uc($AoA[$i][ACCT_CAR_INDEX]);
			$EoE[$arrayIndex][ACCT_LASTPAYMENT_INDEX]       = $AoA[$i][ACCT_LASTPAYMENT_INDEX];
			$EoE[$arrayIndex][ACCT_NAME_INDEX]              = $AoA[$i][ACCT_NAME_INDEX]; 
			$EoE[$arrayIndex][ACCT_REPO_INDEX]              = $AoA[$i][ACCT_REPO_INDEX];
			$EoE[$arrayIndex][ACCT_INSEXPIRE_INDEX]         = $AoA[$i][ACCT_INSEXPIRE_INDEX];
            $EoE[$arrayIndex][ACCT_INSEXPIRE_DELTA_INDEX]   = $AoA[$i][ACCT_INSEXPIRE_DELTA_INDEX];	
			$EoE[$arrayIndex][ACCT_LASTPAYMENT_DELTA_INDEX] = $AoA[$i][ACCT_LASTPAYMENT_DELTA_INDEX];
			$EoE[$arrayIndex][ACCT_SCORE_INDEX]             = $AoA[$i][ACCT_SCORE_INDEX];
			$EoE[$arrayIndex][ACCT_SALE_DATE_DELTA_INDEX]   = $AoA[$i][ACCT_SALE_DATE_DELTA_INDEX];
			$EoE[$arrayIndex][ACCT_PAYOFF_INDEX]            = $AoA[$i][ACCT_PAYOFF_INDEX];				
			$arrayIndex++;
			next;
		}
	} 
}

print "SIZE OF EoE (Sorted by Payment Days Late): ",$#EoE,"\n";


############################################
#FoF is sorted by DAYS LATE COMBO
############################################

for my $i (0 .. $#AoA) 
{
	for my $j (0 .. MAX_INDEX_VALUE-1)
	{
		$FoF[$i][$j] = $AoA[$i][$j];
	}	
}

my $changeMade = 1;

while ($changeMade eq 1)
{
	$changeMade = 0;
	
	for my $i (0 .. ($#FoF - 1)) 
	{
		my $currentIndex = $i;
		my $nextIndex  = $i+1;
		
		my $score     = $FoF[$currentIndex][ACCT_LATE_COMBO_INDEX];		
		my $scoreNext = $FoF[$nextIndex][ACCT_LATE_COMBO_INDEX];
		
		if ( $scoreNext < $score )
		{
			my @ArrayCopy;  
			
			for my $j (0 .. MAX_INDEX_VALUE-1)
			{
				$ArrayCopy[$j] = $FoF[$i][$j];
			}
			
			for my $j (0 .. MAX_INDEX_VALUE-1)
			{
				$FoF[$i][$j] = $FoF[$nextIndex][$j];
			}
			
			for my $j (0 .. MAX_INDEX_VALUE-1)
			{
				$FoF[$nextIndex][$j] = $ArrayCopy[$j];
			}
			
			$changeMade = 1;
			$numChangesMade++;
		}
	}
}	

print "SIZE OF FoF (Sorted by Days Late Combination): ",$#FoF,"\n";


################################
# Print out SCORE REPORT
################################

$myOpenAccountIndex = 1;
$numAccountsSkipped = 0;
print $myScoreOutput "<table id=\"t01\" sytle=width:100%><tr><th>Index</th><th>Score</th><th>Days Late</th><th>Last Payment</th><th>Name</th><th>Vehicle</th><th>Payoff</th><th>Insurance Expiration</th></tr>\n";

for my $i (reverse 0 .. $#BoB)
{
	my $daysLate     	= $BoB[$i][ACCT_DAYSLATE_INDEX];
	my $vehicle      	= $BoB[$i][ACCT_CAR_INDEX];
	my $lastPayment  	= $BoB[$i][ACCT_LASTPAYMENT_INDEX];
	my $lastPaymentDelta = $BoB[$i][ACCT_LASTPAYMENT_DELTA_INDEX];
	my $custname     	= $BoB[$i][ACCT_NAME_INDEX];
	my $insExpire    	= $BoB[$i][ACCT_INSEXPIRE_INDEX];
	my $insDelta     	= $BoB[$i][ACCT_INSEXPIRE_DELTA_INDEX];
	my $score        	= $BoB[$i][ACCT_SCORE_INDEX];
	my $saleDateDelta   = $BoB[$i][ACCT_SALE_DATE_DELTA_INDEX];
	my $payoff          = $BoB[$i][ACCT_PAYOFF_INDEX];
	my $state           = $BoB[$i][ACCT_REPO_INDEX];
	
	if ( $BoB[$i][ACCT_INSEXPIRE_DELTA_INDEX] >= $INSEXP_SKIP )
	{
		$numAccountsSkipped++;
		next;
	}
		
	if (($state ne $REPOSSESSED) && ($state ne $INPROCESSOFREPO) && ($state ne $ONHOLD) )
	{
		my $myStyle;
		
		if (($saleDateDelta <= 90) && ($daysLate >= 15) ) 
		{
			$myStyle = "style=\"background-color: yellow;\"";
		}
		
		print $myScoreOutput "<tr ",$myStyle,"><td>",$myOpenAccountIndex,"</td><td>",$score,"</td><td>",$daysLate,"</td><td>";
		
		if ( $lastPaymentDelta > 30 )
		{
			print $myScoreOutput "<font color=red>",$lastPayment,"</font></td><td>",$custname,"</td><td>",$vehicle,"</td><td>",money_format($payoff),"</td><td>";
		}
		else
		{
			print $myScoreOutput $lastPayment,"</td><td>",$custname,"</td><td>",$vehicle,"</td><td>",money_format($payoff),"</td><td>";
		}
		if ($insDelta > 1 )
		{
			print $myScoreOutput "<font color=red>",$insExpire,"</font></td></tr>\n";
		}
		else
		{
			print $myScoreOutput $insExpire,"</td></tr>\n";
		}
		$myOpenAccountIndex++;
		next;
	}	
}

print $myScoreOutput "</table>\n";
print $myScoreOutput "<br><br><b>Number of Accounts Skipped (Due to Insurance Expiration Dates) = </b>",$numAccountsSkipped,"<br><br>\n";
print $myScoreOutput "<br><br>\n";
print $myScoreOutput "</body></html>\n";

# END - Print out SCORE SORT


################################
# Print out LAST PAYMENT REPORT
################################

$myOpenAccountIndex = 1;
print $myLastPayOutput "<table id=\"t01\" sytle=width:100%><tr><th>Index</th><th>Score</th><th>Days Late</th><th>Last Payment</th><th>Name</th><th>Vehicle</th><th>Insurance Expiration</th></tr>\n";

for my $i (0 .. $#CoC)
{
	my $daysLate     = $CoC[$i][ACCT_DAYSLATE_INDEX];
	my $vehicle      = $CoC[$i][ACCT_CAR_INDEX];
	my $lastPayment  = $CoC[$i][ACCT_LASTPAYMENT_INDEX];
	my $lastPaymentDelta = $CoC[$i][ACCT_LASTPAYMENT_DELTA_INDEX];	
	my $custname     = $CoC[$i][ACCT_NAME_INDEX];
	my $insExpire    = $CoC[$i][ACCT_INSEXPIRE_INDEX];
	my $insDelta     = $CoC[$i][ACCT_INSEXPIRE_DELTA_INDEX];	
	my $score        = $CoC[$i][ACCT_SCORE_INDEX];
	my $state        = $CoC[$i][ACCT_REPO_INDEX];
	
	if (($state ne $REPOSSESSED) && ($state ne $INPROCESSOFREPO) && ($state ne $ONHOLD) && ($daysLate > 0))
	{
		print $myLastPayOutput "<tr><td>",$myOpenAccountIndex,"</td><td>",$score,"</td><td>",$daysLate,"</td><td>";
		
		if ( $lastPaymentDelta > 30 )
		{
			print $myLastPayOutput "<font color=red>",$lastPayment,"</font></td><td>",$custname,"</td><td>",$vehicle,"</td><td>";
		}
		else
		{
			print $myLastPayOutput $lastPayment,"</td><td>",$custname,"</td><td>",$vehicle,"</td><td>";
		}
		if ($insDelta > 1 )
		{
			print $myLastPayOutput "<font color=red>",$insExpire,"</font></td></tr>\n";
		}
		else
		{
			print $myLastPayOutput $insExpire,"</td></tr>\n";
		}
		$myOpenAccountIndex++;
		next;
	}	
}
print $myLastPayOutput "</table>\n";
print $myLastPayOutput "<br><br>\n";
print $myLastPayOutput "</body></html>\n";
# END - Print out LAST PAYMENT SORT

################################
# Print out INSURANCE EXPIRATION
################################

$myOpenAccountIndex = 1;
print $myInsOutput "<table id=\"t01\" sytle=width:100%><tr><th>Index</th><th>Score</th><th>Days Late</th><th>Last Payment</th><th>Name</th><th>Vehicle</th><th>Insurance Expiration</th></tr>\n";
	
for my $i (0 .. $#DoD)
{
	my $daysLate     = $DoD[$i][ACCT_DAYSLATE_INDEX];
	my $vehicle      = uc($DoD[$i][ACCT_CAR_INDEX]);
	my $lastPayment  = $DoD[$i][ACCT_LASTPAYMENT_INDEX];
	my $lastPaymentDelta = $DoD[$i][ACCT_LASTPAYMENT_DELTA_INDEX];	
	my $custname     = $DoD[$i][ACCT_NAME_INDEX];
	my $insExpire    = $DoD[$i][ACCT_INSEXPIRE_INDEX];
	my $insDelta     = $DoD[$i][ACCT_INSEXPIRE_DELTA_INDEX];	
	my $score        = $DoD[$i][ACCT_SCORE_INDEX];	
	my $state        = $DoD[$i][ACCT_REPO_INDEX];
	
	if (($state ne $REPOSSESSED) && ($state ne $INPROCESSOFREPO) && ($state ne $ONHOLD) && ($insDelta > 0))
	{
		print $myInsOutput "<tr><td>",$myOpenAccountIndex,"</td><td>",$score,"</td><td>",$daysLate,"</td><td>";
		
		if ( $lastPaymentDelta > 30 )
		{
			print $myInsOutput "<font color=red>",$lastPayment,"</font></td><td>",$custname,"</td><td>",$vehicle,"</td><td>";
		}
		else
		{
			print $myInsOutput $lastPayment,"</td><td>",$custname,"</td><td>",$vehicle,"</td><td>";
		}
		if ($insDelta > 1 )
		{
			print $myInsOutput "<font color=red>",$insExpire,"</font></td></tr>\n";
		}
		else
		{
			print $myInsOutput $insExpire,"</td></tr>\n";
		}
		$myOpenAccountIndex++;
		next;
	}	
}
print $myInsOutput "</table>\n";
print $myInsOutput "<br><br>\n";
print $myInsOutput "</body></html>\n";
# END - Print out INSURANCE EXPIRATION



################################
# Print out PAYMENT DAYS LATE
################################

$myOpenAccountIndex = 1;
print $myDaysLateOutput "<table id=\"t01\" sytle=width:100%><tr><th>Index</th><th>Score</th><th>Days Late</th><th>Last Payment</th><th>Name</th><th>Vehicle</th><th>Insurance Expiration</th></tr>\n";
	
for my $i (0 .. $#EoE)
{
	my $daysLate     = $EoE[$i][ACCT_DAYSLATE_INDEX];
	my $vehicle      = uc($EoE[$i][ACCT_CAR_INDEX]);
	my $lastPayment  = $EoE[$i][ACCT_LASTPAYMENT_INDEX];
	my $lastPaymentDelta = $EoE[$i][ACCT_LASTPAYMENT_DELTA_INDEX];	
	my $custname     = $EoE[$i][ACCT_NAME_INDEX];
	my $insExpire    = $EoE[$i][ACCT_INSEXPIRE_INDEX];
	my $insDelta     = $EoE[$i][ACCT_INSEXPIRE_DELTA_INDEX];	
	my $score        = $EoE[$i][ACCT_SCORE_INDEX];
	my $state        = $EoE[$i][ACCT_REPO_INDEX];
	
	if (($state ne $REPOSSESSED) && ($state ne $INPROCESSOFREPO) && ($state ne $ONHOLD) && ($daysLate > 0))
	{
		print $myDaysLateOutput "<tr><td>",$myOpenAccountIndex,"</td><td>",$score,"</td><td>",$daysLate,"</td><td>";
		
		if ( $daysLate > 30 )
		{
			print $myDaysLateOutput "<font color=red>",$lastPayment,"</font></td><td>",$custname,"</td><td>",$vehicle,"</td><td>";
		}
		else
		{
			print $myDaysLateOutput $lastPayment,"</td><td>",$custname,"</td><td>",$vehicle,"</td><td>";
		}
		if ($insDelta > 1 )
		{
			print $myDaysLateOutput "<font color=red>",$insExpire,"</font></td></tr>\n";
		}
		else
		{
			print $myDaysLateOutput $insExpire,"</td></tr>\n";
		}
		$myOpenAccountIndex++;
		next;
	}	
}
print $myDaysLateOutput "</table>\n";
print $myDaysLateOutput "<br><br>\n";
print $myDaysLateOutput "</body></html>\n";
# END - Print out PAYMENT DAYS LATE



################################
# Print out LATE ACCOUNTS REPORT
################################

$myOpenAccountIndex = 1;
$numAccountsSkipped = 0;
print $myLateComboOutput "<table id=\"t01\" sytle=width:100%><tr><th>Index</th><th>Days Active</th><th>Days Late</th><th>Last Payment</th><th>Name</th><th>Vehicle</th><th>Insurance Exp.</th><th>Payments Due</th><th>Payoff</th><th>Cell Phone</th><th>Home Phone</th><th>Work Phone</th></tr>\n";
	
for my $i (reverse 0 .. $#FoF)
{
	my $daysLate     = $FoF[$i][ACCT_DAYSLATE_INDEX];
	my $vehicle      = uc($FoF[$i][ACCT_CAR_INDEX]);
	my $lastPayment  = $FoF[$i][ACCT_LASTPAYMENT_INDEX];
	my $lastPaymentDelta = $FoF[$i][ACCT_LASTPAYMENT_DELTA_INDEX];	
	my $custname     = $FoF[$i][ACCT_NAME_INDEX];
	my $insExpire    = $FoF[$i][ACCT_INSEXPIRE_INDEX];
	my $insDelta     = $FoF[$i][ACCT_INSEXPIRE_DELTA_INDEX];	
	my $score        = $FoF[$i][ACCT_SCORE_INDEX];
	my $lateCombo    = $FoF[$i][ACCT_LATE_COMBO_INDEX];
	my $saleDateDelta = $FoF[$i][ACCT_SALE_DATE_DELTA_INDEX];
	my $homephone     = $FoF[$i][ACCT_HOME_PHONE_INDEX];
	my $workphone     = $FoF[$i][ACCT_WORK_PHONE_INDEX];
	my $cellphone     = $FoF[$i][ACCT_CELL_PHONE_INDEX];
	my $payoff        = $FoF[$i][ACCT_PAYOFF_INDEX];
	my $state         = $FoF[$i][ACCT_REPO_INDEX];
    my $totaldue      = $FoF[$i][ACCT_TOTAL_DUE_INDEX];
	my $paymentsdue   = $FoF[$i][ACCT_PAYMENTS_DUE_INDEX];
	
	if ( $FoF[$i][ACCT_INSEXPIRE_DELTA_INDEX] >= $INSEXP_SKIP )
	{
		$numAccountsSkipped++;
		next;
	}
	
	if (($state ne $REPOSSESSED) && ($state ne $INPROCESSOFREPO) && ($state ne $ONHOLD) && ($daysLate > 0))
	{
		my $myStyle;
		
		if (($saleDateDelta <= 90) && ($daysLate >= 15) ) 
		{
			$myStyle = "style=\"background-color: yellow;\"";
		}
		
		print $myLateComboOutput "<tr ",$myStyle,"><td>",$myOpenAccountIndex,"</td><td>",$saleDateDelta,"</td><td>",$daysLate,"</td><td>";
				
		if ( $daysLate > 30 )
		{
			print $myLateComboOutput "<font color=red>",$lastPayment,"</font></td><td>",$custname,"</td><td>",$vehicle,"</td><td>";
		}
		else
		{
			print $myLateComboOutput $lastPayment,"</td><td>",$custname,"</td><td>",$vehicle,"</td><td>";
		}
		if ($insDelta > 1 )
		{
			print $myLateComboOutput "<font color=red>",$insExpire,"</font></td>\n";
		}
		else
		{
			print $myLateComboOutput $insExpire,"</td>\n";
		}
		
		print $myLateComboOutput "<td>",money_format($paymentsdue),"</td>\n";
		print $myLateComboOutput "<td>",money_format($payoff),"</td>\n";
		print $myLateComboOutput "<td>",$cellphone,"</td>\n";
		print $myLateComboOutput "<td>",$homephone,"</td>\n";
		print $myLateComboOutput "<td>",$workphone,"</td></tr>\n";		
		
		$myOpenAccountIndex++;
		next;
	}	
}
print $myLateComboOutput "</table>\n";
print $myLateComboOutput "<br><br><b>Number of Accounts Skipped (Due to Insurance Expiration Dates) = </b>",$numAccountsSkipped,"<br><br>\n";
print $myLateComboOutput "<br><br>\n";
print $myLateComboOutput "</body></html>\n";
# END - Print out PAYMENT DAYS LATE





################################
# Print out REPOSSESSIONs
################################

$myOpenAccountIndex = 1;
print $myOverviewOutput "<br>\n";
print $myOverviewOutput "<table id=\"t02\"><tr><th colspan=7><I>REPOSSESSIONS and ON-HOLD REPORT</I></th></tr></table>";
print $myOverviewOutput "<table id=\"t01\" sytle=width:100%><tr><th>Index</th><th>State</th><th>Days Active</th><th>Last Payment</th><th>Name</th><th>Vehicle</th></tr>\n";
	
for my $i (0 .. $#AoA)
{
	my $saledateDelta = $AoA[$i][ACCT_SALE_DATE_DELTA_INDEX];
	my $daysLate      = $AoA[$i][ACCT_DAYSLATE_INDEX];
	my $vehicle       = uc($AoA[$i][ACCT_CAR_INDEX]);
	my $lastPayment   = $AoA[$i][ACCT_LASTPAYMENT_INDEX];
	my $custname      = $AoA[$i][ACCT_NAME_INDEX];
	my $insExpire     = $AoA[$i][ACCT_INSEXPIRE_INDEX];
	my $score         = $AoA[$i][ACCT_SCORE_INDEX];
	my $state         = $AoA[$i][ACCT_REPO_INDEX];
	
	if (($state eq $REPOSSESSED) || ($state eq $INPROCESSOFREPO) || ($state eq $ONHOLD))
	{
		print $myOverviewOutput "<tr><td>",$myOpenAccountIndex,"</td><td>",$state,"</td><td>",$saledateDelta,"</td><td>",$lastPayment,"</td><td>",$custname,"</td><td>",$vehicle,"</td></tr>\n";
		$myOpenAccountIndex++;
	}	
}
print $myOverviewOutput "</table>\n";
print $myOverviewOutput "<br><br>\n";
print $myOverviewOutput "</body></html>\n";
# END - Print out REPOSSESSIONS

print $myOverviewOutput "<table id=\"t02\"><tr><th colspan=7>ACCOUNT TOTALS REPORT</th></tr></table>";
print $myOverviewOutput "<table id=\"t01\"><tr><th>Days Late</th><th>Total</th><th>Percentage</th>\n";
print $myOverviewOutput "<tr><td>Repossessed</td><td>",$numAcctsRepossessed,"<td>",nearest(.01,($numAcctsRepossessed/$#AoA)*100),"</td></tr>";
print $myOverviewOutput "<tr><td>On-Hold</td><td>",$numAcctsOnHold,"<td>",nearest(.01,($numAcctsOnHold/$#AoA)*100),"</td></tr>";
print $myOverviewOutput "<tr><td>On-Time (< 10)</td><td>",$numAccts0DaysLate+$numAccts1to9DaysLate,"<td>",nearest(.01,(($numAccts0DaysLate+$numAccts1to9DaysLate)/$#AoA)*100),"</td></tr>";
print $myOverviewOutput "<tr><td>1 to 9 </td><td>",$numAccts1to9DaysLate,"<td>",nearest(.01,($numAccts1to9DaysLate/$#AoA)*100),"</td></tr>";
print $myOverviewOutput "<tr><td>10 to 29 </td><td>",$numAccts10to29DaysLate,"<td>",nearest(.01,($numAccts10to29DaysLate/$#AoA)*100),"</td></tr>";
print $myOverviewOutput "<tr><td>30 to 44</td><td>",$numAccts30to44DaysLate,"<td>",nearest(.01,($numAccts30to44DaysLate/$#AoA)*100),"</td></tr>";
print $myOverviewOutput "<tr><td>45 to 59</td><td>",$numAccts45to59DaysLate,"<td>",nearest(.01,($numAccts45to59DaysLate/$#AoA)*100),"</td></tr>";
print $myOverviewOutput "<tr><td>60 to 90</td><td>",$numAccts60to89DaysLate,"<td>",nearest(.01,($numAccts60to89DaysLate/$#AoA)*100),"</td></tr>";
print $myOverviewOutput "<tr><td>90 to 120</td><td>",$numAccts90to120DaysLate,"<td>",nearest(.01,($numAccts90to120DaysLate/$#AoA)*100),"</td></tr>";
print $myOverviewOutput "<tr><td>120 to 180</td><td>",$numAccts120to180DaysLate,"<td>",nearest(.01,($numAccts120to180DaysLate/$#AoA)*100),"</td></tr>";
print $myOverviewOutput "<tr><td>180 to 365</td><td>",$numAccts180to365DaysLate,"<td>",nearest(.01,($numAccts180to365DaysLate/$#AoA)*100),"</td></tr>";
print $myOverviewOutput "<tr><td>365 or greater</td><td>",$numAccts365DaysLateOrGreater,"<td>",nearest(.01,($numAccts365DaysLateOrGreater/$#AoA)*100),"</td></tr>";
print $myOverviewOutput "<tr><td> Total Open Accounts (minus repos and holds): </td><td>",$totalNumOpenAccounts,"</td></tr>";
print $myOverviewOutput "<tr><td> Total Open Accounts (with repos and holds): </td><td>",$totalNumOpenAccounts + $numAcctsRepossessed + $numAcctsOnHold,"</td></tr>";
print $myOverviewOutput "</table>";

print $myOverviewOutput "<br><br>\n";
print $myOverviewOutput "<table id=\"t02\"><tr><th colspan=7>ACCOUNT LAST PAYMENT TOTALS REPORT</th></tr></table>";
print $myOverviewOutput "<table id=\"t01\"><tr><th>Days Since Last Payment</th><th>Total</th><th>Percentage</th>\n";
print $myOverviewOutput "<tr><td>Repossessed</td><td>",$numAcctsRepossessed,"<td>",nearest(.01,($numAcctsRepossessed/$#AoA)*100),"</td></tr>";
print $myOverviewOutput "<tr><td>On-Hold</td><td>",$numAcctsOnHold,"<td>",nearest(.01,($numAcctsOnHold/$#AoA)*100),"</td></tr>";
print $myOverviewOutput "<tr><td>Less than 30</td><td>",$numAccts0to30DaysSinceLastPayments,"<td>",nearest(.01,($numAccts0to30DaysSinceLastPayments/$#AoA)*100),"</td></tr>";
print $myOverviewOutput "<tr><td>30 to 44</td><td>",$numAccts30to44DaysSinceLastPayment,"<td>",nearest(.01,($numAccts30to44DaysSinceLastPayment/$#AoA)*100),"</td></tr>";
print $myOverviewOutput "<tr><td>45 to 59</td><td>",$numAccts45to59DaysSinceLastPayment,"<td>",nearest(.01,($numAccts45to59DaysSinceLastPayment/$#AoA)*100),"</td></tr>";
print $myOverviewOutput "<tr><td>60 to 90</td><td>",$numAccts60to89DaysSinceLastPayment,"<td>",nearest(.01,($numAccts60to89DaysSinceLastPayment/$#AoA)*100),"</td></tr>";
print $myOverviewOutput "<tr><td>90 to 120</td><td>",$numAccts90to120DaysSinceLastPayment,"<td>",nearest(.01,($numAccts90to120DaysSinceLastPayment/$#AoA)*100),"</td></tr>";
print $myOverviewOutput "<tr><td>120 to 180</td><td>",$numAccts120to180DaysSinceLastPayment,"<td>",nearest(.01,($numAccts120to180DaysSinceLastPayment/$#AoA)*100),"</td></tr>";
print $myOverviewOutput "<tr><td>180 to 365</td><td>",$numAccts180to365DaysSinceLastPayment,"<td>",nearest(.01,($numAccts180to365DaysSinceLastPayment/$#AoA)*100),"</td></tr>";
print $myOverviewOutput "<tr><td>365 or greater</td><td>",$numAccts365DaysLateOrGreaterSinceLastPayment,"<td>",nearest(.01,($numAccts365DaysLateOrGreaterSinceLastPayment/$#AoA)*100),"</td></tr>";
print $myOverviewOutput "</table>";

print $myOverviewOutput "<head>\n";
print $myOverviewOutput "<script type=\"text/javascript\" src=\"https://www.google.com/jsapi\"></script>\n";
print $myOverviewOutput "<script type=\"text/javascript\">\n";
print $myOverviewOutput "google.load(\"visualization\", \"1\", {packages:[\"corechart\"]});\n";
print $myOverviewOutput "google.setOnLoadCallback(drawChart);\n";
print $myOverviewOutput "function drawChart() {";

print $myOverviewOutput "var data = google.visualization.arrayToDataTable([\n";
print $myOverviewOutput "['Accounts', 'Late Status'],\n";
print $myOverviewOutput "['1-9 days',",$numAccts1to9DaysLate,"],\n";
print $myOverviewOutput "['10-29 days',",$numAccts10to29DaysLate,"],\n";
print $myOverviewOutput "['30-44 days',",$numAccts30to44DaysLate,"],\n";
print $myOverviewOutput "['45-59 days',",$numAccts45to59DaysLate,"],\n";
print $myOverviewOutput "['60-89 days',",$numAccts60to89DaysLate,"],\n";
print $myOverviewOutput "['90-120 days',",$numAccts90to120DaysLate,"],\n";
print $myOverviewOutput "['120-180 days',",$numAccts120to180DaysLate,"],\n";
print $myOverviewOutput "['180-365 days',",$numAccts180to365DaysLate,"],\n";
print $myOverviewOutput "['>= 365 days late',",$numAccts365DaysLateOrGreater,"],\n";
print $myOverviewOutput "['Repossessed',",$numAcctsRepossessed,"] ]);\n";
print $myOverviewOutput "['On-Hold',",$numAcctsOnHold,"] ]);\n";

print $myOverviewOutput "var options = {title: 'Late Accounts'};\n";
print $myOverviewOutput "var chart = new google.visualization.PieChart(document.getElementById('piechart'));\n";
print $myOverviewOutput "chart.draw(data, options);     }\n";
print $myOverviewOutput "</script>\n";
print $myOverviewOutput "</head>\n";
print $myOverviewOutput "<body>\n";
print $myOverviewOutput "<div id=\"piechart\" style=\"width: 900px; height: 500px;\"></div>\n";
print $myOverviewOutput "</body>\n";

print $myOverviewOutput "<br><br>";
print $myOverviewOutput "<table id=\"t02\"><tr><th colspan=7>OPEN ACCOUNTS FINANCIALS</th></tr></table>";
print $myOverviewOutput "<table id=\"t01\"><tr><th>Description</th><th>Total</th>\n";
print $myOverviewOutput "<tr><td>Total Monthly Payments</td><td>",money_format($totalMontlyPaymentAmount),"<td>";
print $myOverviewOutput "<tr><td>Total Amount Financed </td><td>",money_format($totalAmountFinanced),"<td>";
print $myOverviewOutput "<tr><td>Total Payoff Amount</td><td>",money_format($totalPayoffAmount),"<td>";
print $myOverviewOutput "<tr><td>Total Due (Payoff+Remaining Interest)</td><td>",money_format($totalAmountDue),"<td>";
print $myOverviewOutput "</table>";
print $myOverviewOutput "<br><br>";

print $myOverviewOutput "</html>\n";

close(HTML_OUTPUT_FILE_SCORE);
close(HTML_OUTPUT_FILE_LASTPAYMENT);
close(HTML_OUTPUT_FILE_INSURANCE);
close(HTML_OUTPUT_FILE_OVERVIEW);
close(HTML_OUTPUT_FILE_COMBO_LATE);
close(HTML_OUTPUT_FILE_DAYSLATE);

sub money_format {
  my $number = sprintf "%.2f", shift @_;
  # Add one comma each time through the do-nothing loop
  1 while $number =~ s/^(-?\d+)(\d\d\d)/$1,$2/;
  # Put the dollar sign in the right place
  $number =~ s/^(-?)/$1\$/;
  $number;
}

sub WriteHtmlTopPage 
{
	my $fileHandle = $_[0];
	my $reportName = $_[1];
	my $reportType = $_[2];
	my $myToday = localtime;
	
	print $fileHandle "<html><body>\n";
	print $fileHandle "<title>",$reportType,"</title>\n";
	print $fileHandle "<head><style>\n";
	print $fileHandle "table, th, td { border: 1px black; border-collapse: collapse}\n";
	print $fileHandle "th, td { padding: 10px; text-align: left }\n";
	print $fileHandle "table#t01 tr:nth-child(even) { background-color: #eee; }\n";
	print $fileHandle "table#t01 tr:nth-child(odd) { background-color:#fff; }\n";
	print $fileHandle "table#t01 th { background-color: #084B8A; color: white; }\n";
	print $fileHandle "table#t02 th { background-color: #66CC66; color: black; }\n";
	print $fileHandle "</style></head>\n";
	print $fileHandle "<style TYPE=\"text/css\">";
	print $fileHandle "<!--\n";
	print $fileHandle "TD{font-family: Arial; font-size: 9pt;}\n";
	print $fileHandle "TH{font-family: Arial; font-size: 9pt;}\n";
	print $fileHandle "--->\n";
	print $fileHandle "</style>\n";
	print $fileHandle "<h2><i>",$reportName,"</i></h2>\n";
	print $fileHandle "<table><tr>";
	print $fileHandle "<th colspan=10><font size=3><a href=",$CWD,"\\",$ROOT_DIR,"\\",$overviewFilename,">Accounts Overview</a></th></font>\n";
	print $fileHandle "<th colspan=10><font size=3><a href=",$CWD,"\\",$ROOT_DIR,"\\",$lateWeightedFilename,">Late Accounts</a></th></font>\n";
	print $fileHandle "<th colspan=10><font size=3><a href=",$CWD,"\\",$ROOT_DIR,"\\",$scoreFilename,">Account Scores</a></th></font>\n";
	print $fileHandle "</tr></table>";	
	print $fileHandle "<br>";
	print $fileHandle "<table id=\"t02\"><tr><th colspan=7>",$reportType,"</th><th>",$myToday,"</th><th></th></tr></table>";


}
