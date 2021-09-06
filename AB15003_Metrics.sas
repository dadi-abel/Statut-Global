/*<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
> STUDY         AB15003 tracker           													 				  
>																							
> AUTHOR        Dadi Abel DIEDHIOU														
>																							
> VERSION       n°1																			
> DATE          20210818																	
/*<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

/*TO UPDATE to declare library of SAS DATASETS*/
/*>>>>>>>>>>>>>ALL DATA EXPORTED in SAS >>>>>*/ libname AB15003 "H:\AB15003\01_Data Bases\01_LS_Prod" access=readonly; 
option fmtsearch=(AB15003);
libname Metrics "H:\AB15003\01_Data Bases\Efficacy_parameters";
option fmtsearch =(Metrics);

/*TO UPDATE FOR ENROLL : report from ARISGLOBAL*/
/*>>>>>>>>>>Enrollement status report >>>>>>>*/libname xls excel "H:\AB15003\01_Data Bases\01_LS_Prod\Enrollment-Status-Report.xlsx"  ;
DATA ENROLL;
	SET xls."Enrollment Status Report$"n;
RUN; 
libname xls clear;

/*>>>>>>>>>>Queries report >>>>>>>*/libname xls excel "H:\AB15003\01_Data Bases\01_LS_Prod\Query-Listings-Report.xlsx"  ;
/*TO UPDATE FOR QUERIES : report from ARISGLOBAL*/
DATA DQUERY;
	SET xls."Query Listings Report$"n;
RUN; 
libname xls clear;

/*TO UPDATE FOR MONITORING : report from ARISGLOBAL*/
/*>>>>>>>>>>Current page status report >>>>>>>*/libname xls excel "H:\AB15003\01_Data Bases\01_LS_Prod\CurrentPageStatusReport.xlsx" ;
DATA CMONI;
	SET xls."CurrentPageStatusReport$"n;
RUN; 
libname xls clear;


/*TO UPDATE FOR IWRS: report from ARISGLOBAL*/

PROC IMPORT DATAFILE="H:\AB15003\01_Data Bases\02_IWRS\AB15003_Patients information.xls"  OUT=IWRS  REPLACE
DBMS= xls;
GETNAMES=YES;
RUN;

/*JOIN IWRS AND */
PROC SQL;
CREATE TABLE IWRS1 AS 
SELECT b.Site,a.Patient_number,a.Visit,a.Findings_name
FROM 
IWRS AS A 
FULL JOIN ENROLL AS B 
ON a.Patient_number = b.Screening_ID_
WHERE Visit = "RANDOMISATION" AND Findings_name="posology"
ORDER BY Patient_number
;
QUIT;

  /***********************************************************************************************/
 /*      						Import enroll data and declare library 							*/
/***********************************************************************************************/
/*Create table to vizualize key recruitement data*/
proc sql noprint;
	create table SCREENED as
	select  Site,Location,Screening_ID_, status, Enrollment_Date_Time
	from ENROLL;
quit;

  /*******************************************************************/
 /*					Number screened patient							*/
/*******************************************************************/
proc sql ;
	create table scr_patient as
	select distinct Site,count(status) as Number_scr_patient
	from SCREENED
	group by Site
	having status in ('In study','New');
run;

/* calculate the total of the Number of screened patient*/
proc sql;
create table TOTAL_SCR_PATIENT as
select sum(Number_scr_patient) as Number_scr_patient
 from scr_patient
;
quit;

DATA TOTAL_SCR_PATIENT1;
SET TOTAL_SCR_PATIENT;
Site = "Total";
RUN;

/*Join*/
DATA scr_patient1;
SET scr_patient TOTAL_SCR_PATIENT1;
RUN;

/**/
proc sort data=Ab15003.Rep_studymed_admin out = Rep_studymed_admin;
	by SUBJID;
run;

data test;
set Rep_studymed_admin;
if ECENDAT EQ '' then delete; /*delete the patient without end of date*/
run;


data Rep_studymed_admin1;
set TEST;
if last.SUBJID /*then ECENDAT1 = ECENDAT*/; 
/*keep SUBJID ECENDAT; */
by SUBJID;
run; 

data Rep_studydisc;
set Ab15003.Rep_studydisc;
run;
proc sql;
create table Rep_studydisc1 as 
select Rep_studydisc.*, Rep_studymed_admin1.ECENDAT
from Rep_studydisc AS a left JOIN Rep_studymed_admin1 AS b
ON a.SUBJID=b.SUBJID;
quit;

/*Create table to vizualize key recruitement data (randomized / ongoing / discontinued patient*/
proc sql noprint;
	create table BRANDO as
	select c.*, d.STUDYCOMYN , d.DSTERM, d.ECENDAT , d.DISCONTINUATION_REASON,d.SCRFLRYN
	from
	(select  a.Site, a.Location, a.Screening_ID_, a.status, a.Enrollment_Date_Time, b.DSSTDAT
	from SCREENED a
	left join AB15003.Rep_randomization b
	on a.Screening_ID_=b.SUBJID) c 
	LEFT JOIN 
	(select SUBJID as DSUBJID,STUDYCOMYN, DSTERM, ECENDAT,SCRFLRYN,
					case 		
					 			when R1DSDECOD <> '' then R1DSDECOD
								when R2DSDECOD <> '' then R2DSDECOD
								when R3DSDECOD <> '' then R3DSDECOD
								when R5DSDECOD <> '' then R5DSDECOD
					end as DISCONTINUATION_REASON
	from Rep_studydisc1) d
	on c.Screening_ID_= d.DSUBJID
	;
quit;
/**/
proc sql noprint;
	create table BRANDO_D as
	select Site, Location, Screening_ID_, Status , Enrollment_Date_Time, DSSTDAT, STUDYCOMYN, DSTERM, ECENDAT,
	case 
							when DISCONTINUATION_REASON='' and Status='Cancelled before enrollment' then 'Cancelled before enrollment'
							/*when DISCONTINUATION_REASON='' and Status in ('In study','New') then 'ONGOING'*/
							when STUDYCOMYN='Yes' and Status in ('In study',/*'New'*/) then 'COMPLETED'
							Else DISCONTINUATION_REASON end as DISCONTINUATION_REASON
	from BRANDO a
	;
quit;

/***	RECUPERER LES DATES DE SCREENING DES PATIENTS	***/
DATA REP_VD;
SET AB15003.REP_VD;
where SUBJEVENTNAME ="Screening";
RUN;
PROC SORT DATA=REP_VD; BY SUBJID; RUN;

DATA REP_RANDOMIZATION;
SET AB15003.REP_RANDOMIZATION;
RUN;
PROC SORT DATA=REP_RANDOMIZATION; BY SUBJID; RUN;

/***	FAIRE UNE JOINTURE ENTRE REP_VD ET REP_RANDO	***/
PROC SQL;
CREATE TABLE REP_VD_RANDO AS
SELECT A.*, B.DSSTDAT
FROM REP_VD AS A
LEFT JOIN REP_RANDOMIZATION AS B
ON A.SUBJID = B.SUBJID;
QUIT;

				  /*******************************************************************/
				 /*					Number screen failed patient					*/
				/*******************************************************************/
DATA REP_VD_RANDO1;
SET REP_VD_RANDO;
VSDAT1= INPUT(COMPRESS(VSDAT, "-"),date9.);
FORMAT VSDAT1 DATE9.;
Diff = intck('day', VSDAT1, today());
WHERE DSSTDAT ='';
RUN;

/*MAKE A JOIN BETWEEN REP_VD_RANDO IWRS AND BRANDO*/
PROC SQL ;
CREATE TABLE REP_VD_RANDO2 AS
SELECT  DISTINCT A.STUDYID, A.COUNTRYID, C.Site, A.SUBJID, B.Patient_number, A.VISPER, A.VSDAT1,
A.DSSTDAT, A.Diff
FROM REP_VD_RANDO1 AS A
LEFT JOIN IWRS AS B
ON A.SUBJID = B.Patient_number
LEFT JOIN BRANDO AS C
ON A.SUBJID = C.Screening_ID_;
quit;

/*SCREEN FAILURE PATIENT DETECTION */
DATA BRANDO1 (KEEP = Site SUBJID Patient_number VSDAT1 Diff SCR_FAIL);
SET REP_VD_RANDO2 ;
IF SUBJID NE Patient_number
AND Diff > 16
THEN SCR_FAIL = 'Yes';
ELSE SCR_FAIL = 'No';
LABEL SCR_FAIL = 'Detected a Potential Screen Failure';
LABEL Patient_number = 'Patient Number (IWRS)';
LABEL Diff = 'Number Of Day Between Date Of visit Of Screening and Today';
LABEL SUBJID = 'Patient Number (eCRF)';
LABEL VSDAT1 = 'Date Of visit Of Screening';
RUN;

/*COUNT THE NUMBER OF SCREEN FAILURE PATIENT*/
proc sql ;
	create table scr_failed_patient as 
	select distinct Site, count(SCR_FAIL) as Detected_Screen_Failed_Patient
	from BRANDO1
	group by Site;
run;
/*calculate the total of Number Screen Failed Patient */ 
proc sql ;
create table Total_SCR_FAILED_PATIENT as
select sum(Detected_Screen_Failed_Patient) as Detected_Screen_Failed_Patient
from scr_failed_patient
;
quit;

DATA Total_SCR_FAILED_PATIENT1;
SET  Total_SCR_FAILED_PATIENT;
site = "Total";
RUN;

/*add Total to Screen FAILED PATIENT */
DATA scr_failed_patient1;
Set scr_failed_patient Total_SCR_FAILED_PATIENT1;
LABEL Detected_Screen_Failed_Patient = 'Number Of Detected Screen Failed Patient';
Run;

data Metrics.Screen_Fail;
set BRANDO1;
where SCR_FAIL = 'Yes';
run;

				  /*******************************************************************/
				 /*			   Number randomized patient eCRF				   		*/
				/*******************************************************************/
proc sql ;
	create table rando_patient_eCRF as
	select distinct Site,count(DSSTDAT) as Number_rando_patient_eCRF
	from BRANDO
	group by Site
	having DSSTDAT <> '';
run;
/***	calculate the total of Number of randomized patient **	*/ 
proc sql ;
create table Total_Rando_PATIENT_eCRF as
select sum(Number_rando_patient_eCRF) as Number_rando_patient_eCRF
from rando_patient_eCRF
;
quit;
DATA Total_Rando_PATIENT_eCRF1;
Set  Total_Rando_PATIENT_eCRF;
site = "Total";
RUN;
/***	add Total to Randomized PATIENT 	***/
DATA rando_patient_eCRF1;
Set rando_patient_eCRF Total_Rando_PATIENT_eCRF1;
BY Site;
Run;

			  /*******************************************************************/
			 /*				Number randomized patient IWRS   					*/
			/*******************************************************************/
proc sql ;
	create table rando_patient_iwrs as
	select distinct Site ,count(Visit) as Number_rando_patient_iwrs
	from IWRS1
	group by Site;
run;
/***	calculate the total of Number of randomized patient in IWRS		***/ 
proc sql ;
create table Total_Rando_PATIENT_iwrs as
select sum(Number_rando_patient_iwrs) as Number_rando_patient_iwrs
from rando_patient_iwrs
;
quit;
DATA Total_Rando_PATIENT_IWRS1;
Set  Total_Rando_PATIENT_iwrs;
site = "Total";
RUN;
/***	add Total to Randomized PATIENT 	***/
DATA rando_patient_iwrs1;
Set rando_patient_iwrs Total_Rando_PATIENT_IWRS1;
BY Site;
Run;
/***	JOIN TABLE	***/
DATA RANDOMIZED_Patient;
MERGE rando_patient_eCRF1 rando_patient_iwrs1;
BY SITE;
RUN;

/***	REAJUSTER LA POSITION DE LA LIGNE TOTALE	***/
PROC SORT DATA = RANDOMIZED_Patient; BY DESCENDING Number_rando_patient_iwrs ; RUN;

/*
proc sql noprint;
	create table RANDOMIZED_Patient as
	select Site,Number_rando_patient_eCRF,Number_rando_patient_iwrs
	from rando_patient_eCRF1 as a
	FULL JOIN rando_patient_iwrs1 as b
	on a.Site = b.Site 
	order by Number_rando_patient_iwrs;
quit;*/

					  /********************************************************************/
					 /*				Number STUDY COMPLETED  patient						 */
					/********************************************************************/
proc sql ;
	create table study_completed as 
	select distinct Site,count(STUDYCOMYN) as  NumberStudyCompleted
	from BRANDO
	where DSSTDAT <> ''
	and STUDYCOMYN in ('Yes')
	group by site;
run;
/***	calculate the total of STUDY COMPLETED	***/ 
proc sql ;
create table Total_STUDY_COMPLETED as
select sum(NumberStudyCompleted) as NumberStudyCompleted
from study_completed
;
quit;
DATA Total_STUDY_COMPLETED1;
Set  Total_STUDY_COMPLETED;
site = "Total";
RUN;
/***	add Total to STUDY COMPLETED 	***/
DATA study_completed1;
Set study_completed Total_STUDY_COMPLETED1;
Run;

			  /*******************************************************************/
			 /*					Number DISCONTINUED patient						*/
			/*******************************************************************/
DATA BRANDO2;
SET BRANDO;
where DSSTDAT <> ''
and DISCONTINUATION_REASON <> '';
RUN;
proc sql ;
	create table DISCONTINUED_patient as
	select distinct Site,count(DISCONTINUATION_REASON)as Number_of_discontinued
	from BRANDO2
	group by Site
	having DSSTDAT <> ''
	and DISCONTINUATION_REASON <> '';
run;
/***	calculate the total of Discontinued patient		***/ 
proc sql ;
create table Total_DISCONTINUED_patient as
select sum(Number_of_discontinued) as Number_of_discontinued
from DISCONTINUED_patient
;
quit;
DATA Total_DISCONTINUED_patient1;
Set  Total_DISCONTINUED_patient;
site = "Total";
RUN;
/***	add Total to STUDY COMPLETED 	***/
DATA DISCONTINUED_patient1;
Set DISCONTINUED_patient Total_DISCONTINUED_patient1;
Run;
/*data TOTAL;
set TOTAL_SCR_PATIENT1 Total_SCR_FAILED_PATIENT1 Total_Rando_PATIENT1 Total_STUDY_COMPLETED1 Total_DISCONTINUED_patient1;
Run;*/

/*JOIN TABLE*/
proc sql noprint;
	create table Reporting_table as
	select C.*, D.Number_rando_patient_eCRF,D.Number_rando_patient_iwrs, E.NumberStudyCompleted,F.Number_of_discontinued
	from

	(select  a.*, b.Detected_Screen_Failed_Patient
	from scr_patient1  A /*NUMBER 1 : Number screened patient*/
	LEFT JOIN  scr_failed_patient1 B/*Number screen failed patient*/
	on a.Site = b.Site) C 

	LEFT JOIN 

	(select Site as SiteR, Number_rando_patient_eCRF,Number_rando_patient_iwrs
	from RANDOMIZED_Patient) D /*Number randomized patient*/
	on C.Site= D.SiteR

	LEFT JOIN 

	(select Site as SiteC, NumberStudyCompleted
	from study_completed1) E /*Number STUDY COMPLETED  patient*/
	on C.Site= E.SiteC

	LEFT JOIN 

	(select Site as SiteD, Number_of_discontinued
	from DISCONTINUED_patient1) F /*Number DISCONTINUED patient*/
	on C.Site = F.SiteD
	order by Number_scr_patient
	;
quit;

data Reporting_table;
	set Reporting_table;
	if Number_rando_patient_eCRF=. then Number_rando_patient_eCRF = 0; 
	if Detected_Screen_Failed_Patient=. then Detected_Screen_Failed_Patient = 0;
	if Number_rando_patient_iwrs=. then Number_rando_patient_iwrs = 0;
	if NumberStudyCompleted=. then NumberStudyCompleted = 0;
	if Number_of_discontinued=. then Number_of_discontinued = 0;
run;

				  /*******************************************************************/
				 /*				Count the number of page status						*/
				/*******************************************************************/
/*Cleaning the file*/
proc sql;
	create table CMONI2 as
	select study, country, site_id, physician_id, screening_id,
	Event__Visit_name_ as event, page, page_status
	from CMONI;
quit;

			  /*******************************************************************/
			 /*			NUMBER 1 : Number of pages with open queries			*/
			/*******************************************************************/
proc sql noprint;
	select count(*) into: OQ
	from CMONI2
	where Page_Status='With open and answered queries';
run;
%put &OQ;

			  /******************************************************************************************************************/
			 /*		NUMBER 2 : Number of pages with queries answered but not verified (action to be done by DM or CRA)		   */
			/******************************************************************************************************************/
proc sql noprint;
	select count(*) into: ANQ
	from CMONI2
	where Page_Status='With answered queries';
run;
%put &ANQ;

			  /*******************************************************************/
			 /*			NUMBER 3 : Number of pages fully completed				*/
			/*******************************************************************/
proc sql noprint;
	select count(*) into: FQ
	from CMONI2
	where Page_Status='Completed';
run;
%put &FQ;

			  /*******************************************************************/
			 /*			NUMBER 4: Number of pages potentially missing			*/
			/*******************************************************************/
proc sql noprint;
	select count(*) into: PM
	from CMONI2
	where Page_Status='Incomplete / Untouched';
run;
%put &PM;

			  /*******************************************************************/
			 /*				NUMBER 5: Number of pages INCOMPLETE				*/
			/*******************************************************************/
proc sql noprint;
select count(*) into: IP
from CMONI2
where Page_Status='Incomplete';
run;
%put &IP;

data Reporting_table_MONI;
	ATTRIB 	category format=$1000. label="category";
	ATTRIB 	Results format=8. label="Results";

	/*NUMBER 1 : Number of pages with open queries*/
	Results=&OQ;
	category="Number of pages with open queries";
	output;
	/*NUMBER 2: Number of pages with queries answered but not verified (action to be done by DM or CRA)*/
	Results=&ANQ;
	category="Number of pages with queries answered but not verified (action to be done by DM or CRA)";
	output;
	/*NUMBER 3: Number of pages fully completed*/
	Results=&FQ; 
	category="Number of pages fully completed";
	output;
	/*NUMBER 4: Number of pages potentially missing */
	Results=&PM; 
	category="Number of pages potentially missing : Investigator have not yet completed the page 
	OR have filled a wrong page by mistake : Clarification needed from DM or CRA";
	output;
	/*NUMBER 5: Number of pages INCOMPLETE*/
	Results=&IP; 
	category="Number of pages INCOMPLETE";
	output;
run;

				  /***************************************************************/
				 /*							Query status						*/
				/***************************************************************/

				  /*******************************************************************/
				 /*				NUMBER 1 : Number of  open queries   				*/
				/*******************************************************************/
proc sql;
	create table DQUERY2 as
	select site, Location, Subject_Id_, event, page, field, query_type, Status, Query_Comments, Creation_Time,
	Age_Days,input(Age_Days,4.) as Age_Days1, Last_Reply_Time, input(scan(Last_Reply_Time,1,' '),mmddyy10.) format ddmmyy10. as Date_Last_Rep
	from DQUERY;
quit;

Data Open_queries;
set DQUERY2;
where Status = "Open";
run;

DATA Open_queries1;
Set Open_queries;
/*IF LAST.Screening_ID_;TAKE THE LAST ROW OF THE PATIENT*/
BY Subject_Id_;
RUN	;

proc sql ;
	create table Open_quieries1 as
	select distinct Site,count(Status)as Number_of_open_queries
	from Open_queries
	group by Site;
run;
/*calculate the total of open queries*/ 
proc sql ;
create table Total_Open_Queries as
select sum(Number_of_open_queries) as Number_of_open_queries
from Open_quieries1;
quit;
DATA Total_Open_Queries1;
Set  Total_Open_Queries;
site = "Total";
RUN;
/*add Total of OPEN Quieries */
DATA Open_quieries2;
Set Open_quieries1 Total_Open_Queries1;
Run;

				  /*******************************************************************/
				 /*			NUMBER 1Bis : Number of  open queries > than 30 days  		*/
				/*******************************************************************/
Data Open_queries3;
set DQUERY2;
where Status = "Open" and Age_Days1 > 30;
run;

DATA Open_queries_30;
SET Open_queries3;
/*IF LAST.Screening_ID_;TAKE THE LAST ROW OF THE PATIENT*/
BY Subject_Id_;
RUN	;

proc sql ;
	create table Open_quieries_301 as
	select distinct Site,count(Status)as Number_of_open_queries30
	from Open_queries_30
	group by Site;
run;
/*calculate the total of open queries*/ 
proc sql ;
create table Total_Open_Queries_30 as
select sum(Number_of_open_queries30) as Number_of_open_queries30
from Open_quieries_301;
quit;
DATA Total_Open_Queries_301;
Set  Total_Open_Queries_30;
site = "Total";
RUN;
/*add Total of OPEN Quieries */
DATA Open_quieries302;
Set Open_quieries_301 Total_Open_Queries_301;
Run;

					  /*******************************************************************/
					 /*				NUMBER 2 : Number of  answered queries				*/
					/*******************************************************************/
Data Answered_queries;
set DQUERY2;
where Status = "Answered";
run;

DATA Answered_queries1;
SET Answered_queries;
/*IF LAST.Screening_ID_; TAKE THE LAST ROW OF THE PATIENT*/
BY Subject_Id_;
RUN	;

proc sql ;
	create table Answered_queries2 as
	select distinct Site,count(Status)as Number_of_Answered_queries
	from Answered_queries1
	group by Site;
run;
/*calculate the total of Answered queries*/ 
proc sql ;
create table Total_Answered_Queries as
select sum(Number_of_Answered_queries) as Number_of_Answered_queries
from Answered_queries2;
quit;
DATA Total_Answered_Queries1;
Set  Total_Answered_Queries;
site = "Total";
RUN;
/*add Total of Answered Quieries */
DATA Answered_queries2;
Set Answered_queries2 Total_Answered_Queries1;
Run;


					  /*******************************************************************/
					 /*		NUMBER 2Bis : Number of  answered queries > 1 month			*/
					/*******************************************************************/
Data Answered_queries_30;
set DQUERY2;
where Status = "Answered" and intck('day',Date_Last_Rep,today()) > 30;
BY Subject_Id_;
run;

proc sql ;
	create table Answered_queries_30_1 as
	select distinct Site,count(Status)as Number_of_Answered_queries30
	from Answered_queries_30
	group by Site;
run;

/*calculate the total of Answered queries*/ 
proc sql ;
create table Total_Answered_Queries_30 as
select sum(Number_of_Answered_queries30) as Number_of_Answered_queries30
from Answered_queries_30_1;
quit;

DATA Total_Answered_Queries_30_1;
Set  Total_Answered_Queries_30;
site = "Total";
RUN;
/*add Total of Answered Quieries */
DATA Answered_queries_30_2;
Set Answered_queries_30_1 Total_Answered_Queries_30_1;
Run;

				 /******                          ******/
				/**************************************/
proc sql noprint;
	create table Reporting_table_QUERY as
	select a.*,c.Number_of_open_queries30,b.Number_of_Answered_queries,Number_of_Answered_queries30

	from Open_quieries2 a

	LEFT JOIN Answered_queries2 b

	on a.Site=b.Site 

	LEFT JOIN Open_quieries302 c
	on a.Site=c.Site

	LEFT JOIN Answered_queries_30_2 d
	on a.Site=d.Site

	order by Number_of_Answered_queries;
quit;

					  /***************************************************************************/
					 /*				 QUERIES NOT VERIFIED BY CRA								*/
					/***************************************************************************/
/*Number of queries Answered*/
proc sql noprint;
	create table Queriesnotverified as
	select Site,Status, Creation_Author,Creation_Author_Role,Page,Query_Comments,Query_Type,Last_Reply_Time
	from DQUERY 
	where Status="Answered";
quit;

Data Queriesnotverified(keep= Site Creation_Author Creation_Author_Role Page Query_Comments Query_Type datdiff);
	set Queriesnotverified;
	date	=  input(scan(Last_Reply_Time,1,' '),mmddyy10.);
	format date ddmmyy10.;
	today=%sysfunc(today());
	format today ddmmyy10.;
	datdiff = intck('day',date, today) ;
	label datdiff = "Number of days since last reply";
run;

/*By role and type of query egal 'Manual alert'*/
proc sql noprint;
	create table Queriesnotverified_bis as
	select distinct Creation_Author, Creation_Author_Role, count(*) as total
	from Queriesnotverified
	where Query_Type = "Manual alert"
	group by Creation_Author; 
quit;

/*By role and type of query egal 'Auto alert'*/	
proc sql;
create table Queriesnotverified1 as
select distinct Creation_Author,Creation_Author_Role, count(*) as total
 from Queriesnotverified
 where query_type = 'Auto alert'
 group by Creation_Author
;
quit;

					  /***************************************************************************/
					 /*									PRURITUS								*/
					/***************************************************************************/
proc sql;
	create table PRURITUS1 as
	select a.Screening_ID_ as SUBJID, a.DSSTDAT, b.SUBJEVENTNAME, 
	case			when SUBJEVENTNAME='Baseline' then '0'
					when SUBJEVENTNAME='Screening' then '-2'
					when SUBJEVENTNAME='Final Visit' then '96' /*IF THERE IS NO DISCONTINUATION DETECTED*/
					when SUBJEVENTNAME like 'W%' then substr(SUBJEVENTNAME,3,2)
					end as ORDER
	/*b.CSPERF*/
	from BRANDO a
	LEFT JOIN AB15003.rep_PRURITUS b
	on a.Screening_ID_=b.SUBJID
	order by Screening_ID_;
quit;

/**/
DATA PRURITUS1 ;
	SET PRURITUS1;
	ORDER_D = ORDER*7;
	DROP ORDER;
RUN;

proc sort data = PRURITUS1; by SUBJID SUBJEVENTNAME ORDER_D; run ; 

data PRURITUS2 ;
	set PRURITUS1;
	DSSTDAT = compress(DSSTDAT,'-');
	if length(DSSTDAT)=8 then
	DSSTDAT1 = cats('0',DSSTDAT);
	else DSSTDAT1 = DSSTDAT;
	attrib DSSTDAT_D format=date9. label="CREATED ON";
	attrib VISIT_D format=date9. label="VISIT EXPECTED";
	DSSTDAT_D = input ( DSSTDAT1, date9.);
	/*ORDER_D = input ( ORDER, BEST8.);*/
	IF DSSTDAT_D ne . then VISIT_D = sum(DSSTDAT_D,ORDER_D);
	else VISIT_D = .;

	/*The week of the final visit is unknown*/
	if SUBJEVENTNAME ='Final Visit' then DSSTDAT_D=.;
	if SUBJEVENTNAME ='Final Visit' then VISIT_D=.;
run;
proc sort data=PRURITUS2 out=PRURITUS2; by SUBJID ORDER_D; run;

/***	DERIVATION POTENTIAL VISIT MISSING ACCORDING TO PRIMARY ENDPONT		***/
proc sql;
	create table PRURITUS3 as
	select c.SUBJID, c.ORDER_D, c.EVENT_VISIT, /*c.CSPERF,*/ c.DATE_ECRF, c.THEORITICAL_DATE, d.DISCONTINUATION_REASON
	from
	(select a.SUBJID, a.SUBJEVENTNAME, b.SUBJEVENTNAME as EVENT_VISIT/*c.CSPERF,*/,a.ORDER_D,  b.VSDAT as DATE_ECRF,
	VISIT_D AS THEORITICAL_DATE
	from PRURITUS2 a
	LEFT JOIN AB15003.rep_vd b
	on a.subjid=b.SUBJID
	and a.SUBJEVENTNAME=b.SUBJEVENTNAME
	) c
	LEFT JOIN
	(select SCREENING_ID_ , DISCONTINUATION_REASON from BRANDO) d
	on c.SUBJID=d.SCREENING_ID_
	order by SUBJID, ORDER_D
	;
quit;

/**/
data PRURITUS4;
	set PRURITUS3;
	by subjid ORDER_D;
	today_date=today();
	format today_date date9.;
	if last.subjid then Flag="last visit" ;
	DATE_ECRF = compress(DATE_ECRF,'-'); /*REMOVE THE ASH*/
	if length(DATE_ECRF)=8 then DATE_ECRF1 = cats('0',DATE_ECRF); /*ADD "0"*/
	else DATE_ECRF1 = DATE_ECRF;
	DATE_ECRF2=input(DATE_ECRF1, date9.);
	format DATE_ECRF2 date9.;
	if flag='last visit' and DISCONTINUATION_REASON eq '' then OUTDATED=today_date - DATE_ECRF2;
run;

proc sql;
	create table PRURITUS5 as
	select SUBJID, ORDER_D, EVENT_VISIT, DATE_ECRF, THEORITICAL_DATE,
	DISCONTINUATION_REASON, FLAG, OUTDATED as OUTDATED_DAYS
	from
	PRURITUS4;
quit;

proc sql;
	create table PRURITUS6 as
	select * from PRURITUS5
	where OUTDATED_DAYS is not null
	and EVENT_VISIT  not in ('Final Visit');
quit;

data PRURITUS7;
set PRURITUS6;
if EVENT_VISIT in ('Baseline','W001','W002','W003','W004','W005','W006','W007','W008','W009','W010','W011','W012','W013',
'W014','W015','W016','W017','W018','W019','W020','W021','W022','W023','W024')
and OUTDATED_DAYS <= 28 then statut = "patient not oudated";
if EVENT_VISIT in ('Baseline','W001','W002','W003','W004','W005','W006','W007','W008','W009','W010','W011','W012','W013',
'W014','W015','W016','W017','W018','W019','W020','W021','W022','W023','W024')
and OUTDATED_DAYS > 28 then statut = "patient oudated";

if EVENT_VISIT = 'Screening'
and OUTDATED_DAYS <= 14 then statut = "patient not oudated";
if EVENT_VISIT = 'Screening'
and OUTDATED_DAYS > 14 then statut = "patient oudated";

if EVENT_VISIT in ('W028','W032','W036','W040','W044','W048','W052','W060','W072','W084','W096')
and OUTDATED_DAYS <= 84 then statut = "patient not oudated";
if EVENT_VISIT in ('W028','W032','W036','W040','W044','W048','W052','W060','W072','W084','W096')
and OUTDATED_DAYS > 84 then statut = "patient oudated";
drop ORDER_D;
run;



			  /***************************************************************************/
			 /*									FLUSHES									*/
			/***************************************************************************/
proc sql;
	create table FLUSHES1 as
	select a.Screening_ID_ as SUBJID, a.DSSTDAT, b.SUBJEVENTNAME, 
	case			when SUBJEVENTNAME='Baseline' then '0'
					when SUBJEVENTNAME='Screening' then '-2'
					when SUBJEVENTNAME='Final Visit' then '96' /*IF THERE IS NO DISCONTINUATION DETECTED*/
					when SUBJEVENTNAME like 'W%' then substr(SUBJEVENTNAME,3,2)
					end as ORDER
	/*b.CSPERF*/
	from BRANDO a
	LEFT JOIN AB15003.Rep_flushes b
	on a.Screening_ID_=b.SUBJID
	order by Screening_ID_;
quit;

/**/
DATA FLUSHES1 ;
	SET FLUSHES1;
	ORDER_D = ORDER*7;
	DROP ORDER;
RUN;

proc sort data = FLUSHES1; by SUBJID SUBJEVENTNAME ORDER_D; run ; 

data FLUSHES2 ;
	set FLUSHES1;
	DSSTDAT = compress(DSSTDAT,'-');
	if length(DSSTDAT)=8 then
	DSSTDAT1 = cats('0',DSSTDAT);
	else DSSTDAT1 = DSSTDAT;
	attrib DSSTDAT_D format=date9. label="CREATED ON";
	attrib VISIT_D format=date9. label="VISIT EXPECTED";
	DSSTDAT_D = input ( DSSTDAT1, date9.);
	/*ORDER_D = input ( ORDER, BEST8.);*/
	IF DSSTDAT_D ne . then VISIT_D = sum(DSSTDAT_D,ORDER_D);
	else VISIT_D = .;

	/*The week of the final visit is unknown*/
	if SUBJEVENTNAME ='Final Visit' then DSSTDAT_D=.;
	if SUBJEVENTNAME ='Final Visit' then VISIT_D=.;
run;
proc sort data=FLUSHES2 out=FLUSHES2; by SUBJID ORDER_D; run;

/*DERIVATION POTENTIAL VISIT MISSING ACCORDING TO PRIMARY ENDPONT*/
proc sql;
	create table FLUSHES3 as
	select c.SUBJID, c.ORDER_D, c.EVENT_VISIT, /*c.CSPERF,*/ c.DATE_ECRF, c.THEORITICAL_DATE, d.DISCONTINUATION_REASON
	from
	(select a.SUBJID, a.SUBJEVENTNAME, b.SUBJEVENTNAME as EVENT_VISIT/*c.CSPERF,*/,a.ORDER_D,  b.VSDAT as DATE_ECRF,
	VISIT_D AS THEORITICAL_DATE
	from FLUSHES2 a
	LEFT JOIN AB15003.rep_vd b
	on a.subjid=b.SUBJID
	and a.SUBJEVENTNAME=b.SUBJEVENTNAME
	) c
	LEFT JOIN
	(select SCREENING_ID_ , DISCONTINUATION_REASON from BRANDO) d
	on c.SUBJID=d.SCREENING_ID_
	order by SUBJID, ORDER_D
	;
quit;

/**/
data FLUSHES4;
	set FLUSHES3;
	by subjid ORDER_D;
	today_date=today();
	format today_date date9.;
	if last.subjid then Flag="last visit" ;
	DATE_ECRF = compress(DATE_ECRF,'-'); /*REMOVE THE ASH*/
	if length(DATE_ECRF)=8 then DATE_ECRF1 = cats('0',DATE_ECRF); /*ADD "0"*/
	else DATE_ECRF1 = DATE_ECRF;
	DATE_ECRF2=input(DATE_ECRF1, date9.);
	format DATE_ECRF2 date9.;
	if flag='last visit' and DISCONTINUATION_REASON eq '' then OUTDATED=today_date - DATE_ECRF2;
run;

proc sql;
	create table FLUSHES5 as
	select SUBJID, ORDER_D, EVENT_VISIT, DATE_ECRF, THEORITICAL_DATE,
	DISCONTINUATION_REASON, FLAG, OUTDATED as OUTDATED_DAYS
	from
	FLUSHES4;
quit;

proc sql;
	create table FLUSHES6 as
	select * from FLUSHES5
	where OUTDATED_DAYS is not null
	and EVENT_VISIT  not in ('Final Visit');
quit;

data FLUSHES7;
set FLUSHES6;
if EVENT_VISIT in ('Baseline','W001','W002','W003','W004','W005','W006','W007','W008','W009','W010','W011','W012','W013',
'W014','W015','W016','W017','W018','W019','W020','W021','W022','W023','W024')
and OUTDATED_DAYS <= 28 then statut = "patient not oudated";
if EVENT_VISIT in ('Baseline','W001','W002','W003','W004','W005','W006','W007','W008','W009','W010','W011','W012','W013',
'W014','W015','W016','W017','W018','W019','W020','W021','W022','W023','W024')
and OUTDATED_DAYS > 28 then statut = "patient oudated";

if EVENT_VISIT = 'Screening'
and OUTDATED_DAYS <= 14 then statut = "patient not oudated";
if EVENT_VISIT = 'Screening'
and OUTDATED_DAYS > 14 then statut = "patient oudated";

if EVENT_VISIT in ('W028','W032','W036','W040','W044','W048','W052','W060','W072','W084','W096')
and OUTDATED_DAYS <= 84 then statut = "patient not oudated";
if EVENT_VISIT in ('W028','W032','W036','W040','W044','W048','W052','W060','W072','W084','W096')
and OUTDATED_DAYS > 84 then statut = "patient oudated";
run;


				  /***********************************************************************/
				 /*									HAMILTON							*/
				/***********************************************************************/
proc sql;
	create table HAMD1 as
	select a.Screening_ID_ as SUBJID, a.DSSTDAT, b.SUBJEVENTNAME, 
	case			when SUBJEVENTNAME='Baseline' then '0'
					when SUBJEVENTNAME='Screening' then '-2'
					when SUBJEVENTNAME='Final Visit' then '96' /*IF THERE IS NO DISCONTINUATION DETECTED*/
					when SUBJEVENTNAME like 'W%' then substr(SUBJEVENTNAME,3,2)
					end as ORDER
	/*b.CSPERF*/
	from BRANDO a
	LEFT JOIN AB15003.Rep_hamd b
	on a.Screening_ID_=b.SUBJID
	order by Screening_ID_;
quit;

/**/
DATA HAMD1 ;
	SET HAMD1;
	ORDER_D = ORDER*7;
	DROP ORDER;
RUN;

proc sort data = HAMD1; by SUBJID SUBJEVENTNAME ORDER_D; run ; 

data HAMD2 ;
	set HAMD1;
	DSSTDAT = compress(DSSTDAT,'-');
	if length(DSSTDAT)=8 then
	DSSTDAT1 = cats('0',DSSTDAT);
	else DSSTDAT1 = DSSTDAT;
	attrib DSSTDAT_D format=date9. label="CREATED ON";
	attrib VISIT_D format=date9. label="VISIT EXPECTED";
	DSSTDAT_D = input ( DSSTDAT1, date9.);
	/*ORDER_D = input ( ORDER, BEST8.);*/
	IF DSSTDAT_D ne . then VISIT_D = sum(DSSTDAT_D,ORDER_D);
	else VISIT_D = .;

	/*The week of the final visit is unknown*/
	if SUBJEVENTNAME ='Final Visit' then DSSTDAT_D=.;
	if SUBJEVENTNAME ='Final Visit' then VISIT_D=.;
run;
proc sort data=HAMD2 out=HAMD2; by SUBJID ORDER_D; run;

/*DERIVATION POTENTIAL VISIT MISSING ACCORDING TO PRIMARY ENDPONT*/
proc sql;
	create table HAMD3 as
	select c.SUBJID, c.ORDER_D, c.EVENT_VISIT, /*c.CSPERF,*/ c.DATE_ECRF, c.THEORITICAL_DATE, d.DISCONTINUATION_REASON
	from
	(select a.SUBJID, a.SUBJEVENTNAME, b.SUBJEVENTNAME as EVENT_VISIT/*c.CSPERF,*/,a.ORDER_D,  b.VSDAT as DATE_ECRF,
	VISIT_D AS THEORITICAL_DATE
	from HAMD2 a
	LEFT JOIN AB15003.rep_vd b
	on a.subjid=b.SUBJID
	and a.SUBJEVENTNAME=b.SUBJEVENTNAME
	) c
	LEFT JOIN
	(select SCREENING_ID_ , DISCONTINUATION_REASON from BRANDO) d
	on c.SUBJID=d.SCREENING_ID_
	order by SUBJID, ORDER_D
	;
quit;

/**/
data HAMD4;
	set HAMD3;
	by subjid ORDER_D;
	today_date=today();
	format today_date date9.;
	if last.subjid then Flag="last visit" ;
	DATE_ECRF = compress(DATE_ECRF,'-'); /*REMOVE THE ASH*/
	if length(DATE_ECRF)=8 then DATE_ECRF1 = cats('0',DATE_ECRF); /*ADD "0"*/
	else DATE_ECRF1 = DATE_ECRF;
	DATE_ECRF2=input(DATE_ECRF1, date9.);
	format DATE_ECRF2 date9.;
	if flag='last visit' and DISCONTINUATION_REASON eq '' then OUTDATED=today_date - DATE_ECRF2;
run;

proc sql;
	create table HAMD5 as
	select SUBJID, ORDER_D, EVENT_VISIT, DATE_ECRF, THEORITICAL_DATE,
	DISCONTINUATION_REASON, FLAG, OUTDATED as OUTDATED_DAYS
	from
	HAMD4;
quit;

proc sql;
	create table HAMD6 as
	select * from HAMD5
	where OUTDATED_DAYS is not null
	and EVENT_VISIT  not in ('Final Visit');
quit;

data HAMD7;
set HAMD6;
if EVENT_VISIT in ('Baseline','W001','W002','W003','W004','W005','W006','W007','W008','W009','W010','W011','W012','W013',
'W014','W015','W016','W017','W018','W019','W020','W021','W022','W023','W024')
and OUTDATED_DAYS <= 28 then statut = "patient not oudated";
if EVENT_VISIT in ('Baseline','W001','W002','W003','W004','W005','W006','W007','W008','W009','W010','W011','W012','W013',
'W014','W015','W016','W017','W018','W019','W020','W021','W022','W023','W024')
and OUTDATED_DAYS > 28 then statut = "patient oudated";

if EVENT_VISIT = 'Screening'
and OUTDATED_DAYS <= 14 then statut = "patient not oudated";
if EVENT_VISIT = 'Screening'
and OUTDATED_DAYS > 14 then statut = "patient oudated";

if EVENT_VISIT in ('W028','W032','W036','W040','W044','W048','W052','W060','W072','W084','W096')
and OUTDATED_DAYS <= 84 then statut = "patient not oudated";
if EVENT_VISIT in ('W028','W032','W036','W040','W044','W048','W052','W060','W072','W084','W096')
and OUTDATED_DAYS > 84 then statut = "patient oudated";
run;


					  /******************************************************************************************/
					 /*									FATIGUE SEVERITY SCALE								   */
					/******************************************************************************************/

proc sql;
	create table FSS1 as
	select a.Screening_ID_ as SUBJID, a.DSSTDAT, b.SUBJEVENTNAME, 
	case			when SUBJEVENTNAME='Baseline' then '0'
					when SUBJEVENTNAME='Screening' then '-2'
					when SUBJEVENTNAME='Final Visit' then '96' /*IF THERE IS NO DISCONTINUATION DETECTED*/
					when SUBJEVENTNAME like 'W%' then substr(SUBJEVENTNAME,3,2)
					end as ORDER
	/*b.CSPERF*/
	from BRANDO a
	LEFT JOIN AB15003.Rep_fss b
	on a.Screening_ID_=b.SUBJID
	order by Screening_ID_;
quit;

/**/
DATA FSS1 ;
	SET FSS1;
	ORDER_D = ORDER*7;
	DROP ORDER;
RUN;

proc sort data = FSS1; by SUBJID SUBJEVENTNAME ORDER_D; run ; 

data FSS2 ;
	set FSS1;
	DSSTDAT = compress(DSSTDAT,'-');
	if length(DSSTDAT)=8 then
	DSSTDAT1 = cats('0',DSSTDAT);
	else DSSTDAT1 = DSSTDAT;
	attrib DSSTDAT_D format=date9. label="CREATED ON";
	attrib VISIT_D format=date9. label="VISIT EXPECTED";
	DSSTDAT_D = input ( DSSTDAT1, date9.);
	/*ORDER_D = input ( ORDER, BEST8.);*/
	IF DSSTDAT_D ne . then VISIT_D = sum(DSSTDAT_D,ORDER_D);
	else VISIT_D = .;

	/*The week of the final visit is unknown*/
	if SUBJEVENTNAME ='Final Visit' then DSSTDAT_D=.;
	if SUBJEVENTNAME ='Final Visit' then VISIT_D=.;
run;
proc sort data=FSS2 out=FSS2; by SUBJID ORDER_D; run;

/*DERIVATION POTENTIAL VISIT MISSING ACCORDING TO PRIMARY ENDPONT*/
proc sql;
	create table FSS3 as
	select c.SUBJID, c.ORDER_D, c.EVENT_VISIT, /*c.CSPERF,*/ c.DATE_ECRF, c.THEORITICAL_DATE, d.DISCONTINUATION_REASON
	from
	(select a.SUBJID, a.SUBJEVENTNAME, b.SUBJEVENTNAME as EVENT_VISIT/*c.CSPERF,*/,a.ORDER_D,  b.VSDAT as DATE_ECRF,
	VISIT_D AS THEORITICAL_DATE
	from FSS2 a
	LEFT JOIN AB15003.rep_vd b
	on a.subjid=b.SUBJID
	and a.SUBJEVENTNAME=b.SUBJEVENTNAME
	) c
	LEFT JOIN
	(select SCREENING_ID_ , DISCONTINUATION_REASON from BRANDO) d
	on c.SUBJID=d.SCREENING_ID_
	order by SUBJID, ORDER_D
	;
quit;

/**/
data FSS4;
	set FSS3;
	by subjid ORDER_D;
	today_date=today();
	format today_date date9.;
	if last.subjid then Flag="last visit" ;
	DATE_ECRF = compress(DATE_ECRF,'-'); /*REMOVE THE ASH*/
	if length(DATE_ECRF)=8 then DATE_ECRF1 = cats('0',DATE_ECRF); /*ADD "0"*/
	else DATE_ECRF1 = DATE_ECRF;
	DATE_ECRF2=input(DATE_ECRF1, date9.);
	format DATE_ECRF2 date9.;
	if flag='last visit' and DISCONTINUATION_REASON eq '' then OUTDATED=today_date - DATE_ECRF2;
run;

proc sql;
	create table FSS5 as
	select SUBJID, ORDER_D, EVENT_VISIT, DATE_ECRF, THEORITICAL_DATE,
	DISCONTINUATION_REASON, FLAG, OUTDATED as OUTDATED_DAYS
	from
	FSS4;
quit;

proc sql;
	create table FSS6 as
	select * from FSS5
	where OUTDATED_DAYS is not null
	and EVENT_VISIT  not in ('Final Visit');
quit;

data FSS7;
set FSS6;
if EVENT_VISIT in ('Baseline','W001','W002','W003','W004','W005','W006','W007','W008','W009','W010','W011','W012','W013',
'W014','W015','W016','W017','W018','W019','W020','W021','W022','W023','W024')
and OUTDATED_DAYS <= 28 then statut = "patient not oudated";
if EVENT_VISIT in ('Baseline','W001','W002','W003','W004','W005','W006','W007','W008','W009','W010','W011','W012','W013',
'W014','W015','W016','W017','W018','W019','W020','W021','W022','W023','W024')
and OUTDATED_DAYS > 28 then statut = "patient oudated";

if EVENT_VISIT = 'Screening'
and OUTDATED_DAYS <= 14 then statut = "patient not oudated";
if EVENT_VISIT = 'Screening'
and OUTDATED_DAYS > 14 then statut = "patient oudated";

if EVENT_VISIT in ('W028','W032','W036','W040','W044','W048','W052','W060','W072','W084','W096')
and OUTDATED_DAYS <= 84 then statut = "patient not oudated";
if EVENT_VISIT in ('W028','W032','W036','W040','W044','W048','W052','W060','W072','W084','W096')
and OUTDATED_DAYS > 84 then statut = "patient oudated";
run;

							  /*******************************************************************/
							 /*									AE								*/
							/*******************************************************************/
Data AE;
	set Ab15003.rep_ae;
run;

/*NUMBER 0 : Number of  AE */
proc sql noprint;
	select count(*) into: AE
	from AE
	;
run;
%put &AE;

/*NUMBER 1 : Number of  AE serious*/
proc sql noprint;
	select count(*) into: AE_SER
	from AE
	where upper(AESER)='YES';
run;
%put &AE_SER;

/*NUMBER 2 : Number of  AE ONGOING*/
proc sql noprint;
	select count(*) into: AE_ONGO
	from AE
	where upper(AEONGO)='YES';
run;
%put &AE_ONGO;

/*Check 1 : Start date after end date / end date before start date*/
data CHECK1 ;
	set AE;
	STARTDATE = compress(AESTDTC,'-');/*Enlever le tiret*/
	if length(STARTDATE)=8 then STARTDATE_D = cats('0',STARTDATE); /*Nouvelle variable avec le '0' du début manquant rajouté*/
	else STARTDATE_D = STARTDATE;
	attrib STARTDATE_DD format=date9. label="CREATED ON";
	STARTDATE_DD = input ( STARTDATE_D, date9.);

	ENDDATE = compress(AEENDTC,'-');/*Enlever le tiret*/
	if length(ENDDATE)=8 then ENDDATE_D = cats('0',ENDDATE); /*Nouvelle variable avec le '0' du début manquant rajouté*/
	else ENDDATE_D = ENDDATE;
	attrib ENDDATE_DD format=date9. label="CREATED ON";
	ENDDATE_DD = input ( ENDDATE_D, date9.);
run;
/*Harmonize date*/
proc sql;
	create table AE2 as
	select SUBJID, SUBJEVENTNAME, AESPID, AETERM, STARTDATE_DD as AESTART, ENDDATE_DD as AEEND, AEONGO, AESER, AEREFID, AESDTH, 
	AESLIFE, AESDISAB, AESHOSP, AESMIE, AESCONG, AESEV, AEINT, AEREL, AERELNST, OTHER_SPECIFY, AEACN, AEACNOTH_NONE, AEACNOTH_CONMED
	from CHECK1;
quit;
/**/
proc sql;
	create table AE4 as
	select * from ae2 a
	LEFT JOIN 
	(select Screening_ID_, DISCONTINUATION_REASON from brando_d) b
	on a.SUBJID=b.Screening_ID_ 
	;
quit;
/**/
proc sql;
	create table AE5 as
	select a.*, 
	case 
				when AEACN='Permanently discontinued' and AEREL='Related to study treatment' and DISCONTINUATION_REASON='Toxicity/AE' then 'Toxicity related and drug suspected' 
				when AESTART is not null and AEEND is not null and (AESTART> AEEND) then  'Start date before end date'
				when AEEND is not null and (AEEND < AESTART) then 'End date before start date'
				when AEEND is not null and AEONGO='Yes' then 'AE has an end date, but ongoing is ticked'
				end as FLAG

	from 
	AE4 a

	;
quit;


			  /******************************************************************************************/
			 /***************				MONITORING STATUS (SDV Status)   		********************/
			/******************************************************************************************/

				/*******			List the number of page entered in the ecrf		  *********/
PROC SQL ;
	CREATE TABLE page_entred AS
	SELECT Site_ID, Page_Status, SDV_Status, Event_Status
	FROM CMONI
	WHERE
	(Page_Status NOT IN ('Completed / Untouched', 'Incomplete / Untouched', 'Not started') )
	and Event_Status NE 'Cancelled'
	ORDER BY Site_ID;
QUIT;


								/*******			Count the number of page entered in the ecrf by site		  *********/
							   /******************************************************************************************/
PROC SQL ;
	CREATE TABLE page_entred_site AS
	SELECT Site_ID,count(*) as Number_Page_Entred
	FROM page_entred
	GROUP BY Site_ID
	ORDER BY Site_ID;
QUIT;

							 /********       calculate the total of page entered in the ecrf		***********/ 
							/*********************************************************************************/
proc sql ;
create table Total_Page_Entered as
select sum(Number_Page_Entred) as Number_Page_Entred
from page_entred_site
;
quit;

DATA Total_Page_Entered1;  
Set  Total_Page_Entered;
Site_ID = "Total";
RUN;
/*** make a join 
DATA Total_Page_Entered2;
SET page_entred_site Total_Page_Entered1;
Run; ***/

								/********	select page status not SDV 	***************/
							   /******************************************************/
PROC SQL ;
	CREATE TABLE Not_SDV_Status AS 
	SELECT Site_ID, count(*) as Number_of_Not_SDV
	FROM page_entred
	WHERE SDV_Status EQ ''
	GROUP BY Site_ID
	ORDER BY Site_ID;
QUIT;

 							 /********       	calculate the total of page status not SDV			***********/ 
							/*********************************************************************************/
proc sql ;
create table Total_Not_SDV_Status as
select sum(Number_of_Not_SDV) as Number_of_Not_SDV
from Not_SDV_Status
;
quit;

DATA Total_Not_SDV_Status1;
Set  Total_Not_SDV_Status;
Site_ID = "Total";
RUN;
/*** make a join 
DATA Total_Not_SDV_Status2;
SET Not_SDV_Status Total_Not_SDV_Status1;
Run; ***/

							/********	 select page status SDV     	*****************/
						   /************************************************************/
PROC SQL ;
	CREATE TABLE SDV_Status AS 
	SELECT Site_ID,count(*) as Number_SDV
	FROM page_entred
	WHERE
	/*Page_Status IN ('Completed', 'Incomplete', 'With open and answered queries', 'With open queries', 'With answered queries' )
	AND*/ SDV_Status NE ''
	/*AND Event_Status IN ('Completed','Incomplete')*/
	GROUP BY Site_ID
	ORDER BY Site_ID;
QUIT;


										 /***	calculate the total of page status not SDV	***/ 
										/*****************************************************/
proc sql ;
create table Total_SDV_Status as
select sum(Number_SDV) as Number_SDV
from SDV_Status
;
quit;

DATA Total_SDV_Status1;
Set  Total_SDV_Status;
Site_ID = "Total";
RUN;
/*** make a join 
DATA Total_SDV_Status2;
SET SDV_Status Total_SDV_Status1;
Run; ***/

				 /*********		Join all table to have one table synthese		***********/
				/*************************************************************************/
data table;
	merge page_entred_site Not_SDV_Status SDV_Status;
	by Site_ID;
run;

/**/
data total;
	merge Total_Page_Entered1 Total_Not_SDV_Status1 Total_SDV_Status1;
	by Site_ID;
run;

data monitoting;
	set table total;
run;
 					 /*************************    RECAP STUDY    **********************/
					/******************************************************************/

libname recap "M:\21 Data Management\DADI\AB19001\Recap Study";

data recap.Reporting_query_AB15003;
set Reporting_table_query;
run;

data recap.Reporting_table_Nb_Page_AB15003;
set Reporting_table_MONI;
run;

data recap.Monitoting_AB15003;
set Monitoting;
run;

data recap.Reporting_table1_Rando_AB15003;
set Reporting_table;
run;



					  /*******************************************************************************************/
					 /*		ADD RECONCILIATION IWRS/ BE INFORME THAT YOU HAVE TO UPDATE THE IWRS FILE BEFOR		*/
					/*******************************************************************************************/

%INCLUDE "H:\AB15003\02_SAS Programs\Metrics\IWRS_RECONCILIATION_AB15003.sas";

				  /*******************************************************************************************/
				 /*		  ADD CALCUL RATIO // BE INFORME THAT YOU HAVE TO UPDATE THE IWRS FILE BEFOR		*/
				/*******************************************************************************************/
 
%INCLUDE "H:\AB15003\02_SAS Programs\Metrics\AB15003_Calcul_Ratio.sas";

			  /*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<*/
			 /*							     QUALITY CONTROL										  */
			/*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<*/

%INCLUDE "H:\AB15003\02_SAS Programs\Metrics\Contrôle qualite.sas";

		  /*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<*/
		 /*							EXPORTATION EXCEL FILE										  */
		/*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<*/

ODS EXCEL FILE = "H:\AB15003\02_SAS Programs\Metrics\Data_Out\AB15003_%sysfunc(date(), date9.).xlsx"
	options( embedded_titles="yes" sheet_name="Monitoring" orientation='landscape');

/*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<*/
ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "Enrollement global status");
option missing = "";
PROC REPORT DATA= brando_d

    nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "All patient global recruitment";

RUN;

/*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<*/
ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "Global table enrollement");
option missing = 0;/*options missing to replace . by 0*/
PROC REPORT DATA= Reporting_table

    nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];

	DEFINE Number_scr_patient / "Number of screened patients";
	DEFINE Number_rando_patient_eCRF / "Number of randomized patients in eCRF";
	DEFINE Number_rando_patient_iwrs / "Number of randomized patients in iwrs";
	DEFINE Number_of_discontinued / "Number of discontinued";

RUN;

/*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<*/
ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "Detection Potential Screen Failure");
option missing = "";
PROC REPORT DATA= BRANDO1

    nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "Detection Potential Screen Failure";

RUN;

/*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<*/
ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "Monitoring global table");
option missing = "";
PROC REPORT DATA= cmoni2

    nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "Monitoring global table";

RUN;

/*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<*/
ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "Reporting global summary");
option missing = "";
PROC REPORT DATA= Reporting_table_moni
    nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "Monitoring global summarize";

RUN;

/*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<*/
ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "Queries global table");
option missing = 0; /*options missing to replace . by 0*/
PROC REPORT DATA= Reporting_table_query

    nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "Query global summarize by site";

	DEFINE Number_of_open_queries / "Number of open queries";
	DEFINE Number_of_open_queries30 / "Number of open queries>30";
	DEFINE Number_of_Answered_queries / "Number of answered queries";

RUN;

/*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<*/
ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "QUERIES TBC");
option missing = "";
PROC REPORT DATA = Queriesnotverified

	nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "QUERIES TBC";

RUN;

/*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<*/
ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "Queries TBC (Manual Queries)";
option missing = "";
PROC REPORT DATA = Queriesnotverified_bis

	nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "Queries TBC (Manual Queries)";

RUN;
/*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<*/
ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "Queries TBC (Automatic Queries)";
option missing = "";
PROC REPORT DATA = Queriesnotverified1

	nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "Queries TBC (Automatic Queries)";

RUN;
/*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<*/
ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "OUTDATED PATIENT(PRURITE)");
option missing = "";
PROC REPORT DATA= Pruritus7

    nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "OUTDATED PATIENT(PRURITE)";

	DEFINE OUTDATED_DAYS / "Number of days between the last visit and today";
	
	COMPUTE statut ;
	IF statut = "patient not oudated" THEN CALL DEFINE (_COL_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTGREEN]");
	ENDCOMP;

RUN;

/*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<*/
ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "OUTDATED PATIENT(FLUSHES)");
option missing = "";
PROC REPORT DATA= Flushes7

    nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "OUTDATED PATIENT(FLUSHES)";

	DEFINE OUTDATED_DAYS / "Number of days between the last visit and today";
	
	COMPUTE statut ;
	IF statut = "patient not oudated" THEN CALL DEFINE (_COL_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTGREEN]");
	ENDCOMP;

RUN;

/*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<*/
ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "OUTDATED PATIENT(HAMILTON)");
option missing = "";
PROC REPORT DATA= Hamd7

    nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "OUTDATED PATIENT(HAMILTON)";

	DEFINE OUTDATED_DAYS / "Number of days between the last visit and today";
	
	COMPUTE statut ;
	IF statut = "patient not oudated" THEN CALL DEFINE (_COL_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTGREEN]");
	ENDCOMP;

RUN;

/*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<*/
ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "OUTDATED PATIENT(FATIGUE SEVERITY SCALE)");
option missing = "";
PROC REPORT DATA= Fss7

    nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "OUTDATED PATIENT(FATIGUE SEVERITY SCALE)";

	DEFINE OUTDATED_DAYS / "Number of days between the last visit and today";
	
	COMPUTE statut ;
	IF statut = "patient not oudated" THEN CALL DEFINE (_COL_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTGREEN]");
	ENDCOMP;

RUN;

/*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<*/
ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "Adverse event check table");
option missing = "";
PROC REPORT DATA= AE5

    nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "OUTDATED PATIENT";


RUN;

/*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<*/
ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "Quality control");
option missing = "";
/*option missing = 0;options missing to replace . by 0*/
PROC REPORT DATA= Rep_vs1

    nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "Quality control";

RUN;

/*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<*/
ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "Quality control");
option missing = "";
/*option missing = 0;options missing to replace . by 0*/
PROC REPORT DATA= Rep_ifc_dm1

    nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "Quality control";

RUN;

/*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<*/
/*ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "SOURCE DATA VERIFICATION (SDV)");
option missing = "";
/*option missing = 0;options missing to replace . by 0*/
/*PROC REPORT DATA= Total_sdv_status2

    nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "SOURCE DATA VERIFICATION (SDV)";

RUN;*/

/*>>>>>>>>>>>>>>>>>>>>>		MONITORING STATUS	<<<<<<<<<<<<<<<<<<<<<*/

ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "MONITORING STATUS");
option missing = "";
PROC REPORT DATA= Monitoting

    nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "MONITORING STATUS";
	
	DEFINE Number_Page_Entred / "Number page entred in the eCRF";
	DEFINE Number_of_Not_SDV / "Number of page still to be SDV";
	DEFINE Number_SDV / "Number of page SDV";

RUN;

/*Close file*/

ODS excel CLOSE;

							  /*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<*/
							 /*									RECONCIALIATION										  */
							/*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<*/

ODS  EXCEL
	FILE = "H:\AB15003\02_SAS Programs\Metrics\Data_Out\AB15003_RECONCILIATION_%sysfunc(date(), date9.).xlsx"
	options( embedded_titles="yes" sheet_name="Monitoring" orientation='landscape');

/*>>>>>>>>>>>>>>>>>>>>>>		SEX			<<<<<<<<<<<<<<<<<<<*/

ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "SEX");
option missing = "";
PROC REPORT DATA = TAB_SEX1

	nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "SEX";

	DEFINE TEST_SEX / "Gender Reconciliation result";
	DEFINE Findings_value / "Gender (IWRS)";
	DEFINE SEX / "Gender (eCRF)";
	
	COMPUTE TEST_SEX ;
	IF TEST_SEX = "COHERENT" THEN CALL DEFINE (_COL_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTGREEN]");
	ELSE IF TEST_SEX = "INCOHERENT" THEN CALL DEFINE (_COL_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTRED]");
	ENDCOMP;

RUN;

					/*>>>>>>>>>>>>>>>>>>>>>>>		WEIGHT		<<<<<<<<<<<<<<<<<<<<*/

ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "WEIGHT");
option missing = "";
PROC REPORT DATA = TAB_WEIGHT1

	nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "WEIGHT";

	DEFINE TEST_WEIGHT / "Weight Reconciliation result";
	DEFINE Findings_value1 / "Weight (IWRS)";
	DEFINE WEIGHT / "Weight (eCRF)";
	
	COMPUTE TEST_WEIGHT ;
	IF TEST_WEIGHT = "COHERENT" THEN CALL DEFINE (_COL_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTGREEN]");
	ELSE IF TEST_WEIGHT = "INCOHERENT" THEN CALL DEFINE (_COL_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTRED]");
	ENDCOMP;

RUN;

					/*>>>>>>>>>>>>>>>>>>>>>>		HEIGHT		<<<<<<<<<<<<<<<<<<<<*/

ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "HEIGHT");
option missing = "";
PROC REPORT DATA = TAB_HEIGHT1

	nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "HEIGHT";

	DEFINE TEST_HEIGHT / "Height Reconciliation result";
	DEFINE Findings_value1 / "Height (IWRS)";
	DEFINE HEIGHT_C / "Height (eCRF)";
	
	COMPUTE TEST_HEIGHT ;
	IF TEST_HEIGHT = "COHERENT" THEN CALL DEFINE (_COL_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTGREEN]");
	ELSE IF TEST_HEIGHT = "INCOHERENT" THEN CALL DEFINE (_COL_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTRED]");
	ENDCOMP;

RUN;

				/*>>>>>>>>>>>>>>>>>>>>>>>>		Date Of Birth		<<<<<<<<<<<<<<<<<<<*/

ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "Date Of Birth");
option missing = "";
PROC REPORT DATA = TAB_BIRTHDATE1

	nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "Date Of Birth";

	DEFINE TEST_HBIRTHDATE / "Date Of Birth Reconciliation result";
	DEFINE Findings_value1 / "Date Of Birth (IWRS)";
	DEFINE Date_Birth / "Date Of Birth (eCRF)";
	DEFINE YEAR_BTH / "Year Of Birth (eCRF)";
	
	COMPUTE TEST_HBIRTHDATE ;
	IF TEST_HBIRTHDATE = "COHERENT" THEN CALL DEFINE (_COL_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTGREEN]");
	ELSE IF TEST_HBIRTHDATE = "INCOHERENT" THEN CALL DEFINE (_COL_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTRED]");
	ENDCOMP;

RUN;

/*>>>>>>>>>>>>>>>>>>>>>>>>			BMI			<<<<<<<<<<<<<<<<<<<*/

ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "BMI");
option missing = "";
PROC REPORT DATA = TAB_BMI_1

	nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "BMI";

	DEFINE TEST_BMI / "BMI Reconciliation result";
	DEFINE Findings_value1 / "BMI (IWRS)";
	DEFINE BMI / "BMI (eCRF)";
	
	COMPUTE TEST_BMI ;
	IF TEST_BMI = "COHERENT" THEN CALL DEFINE (_COL_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTGREEN]");
	ELSE IF TEST_BMI = "INCOHERENT" THEN CALL DEFINE (_COL_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTRED]");
	ENDCOMP;

RUN;

				/*>>>>>>>>>>>>>>>>>>>>>>		RURITUS SCORE		<<<<<<<<<<<<<<<<<<<*/

ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "PRURITUS SCORE");
option missing = "";
PROC REPORT DATA = TAB_PRURITUS1

	nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "PRURITUS SCORE";

	DEFINE TEST_PRURSCOR / "Pruritus Score Reconciliation result";
	DEFINE Findings_value1 / "Pruritus Score (IWRS)";
	DEFINE PRURSCOR / "Pruritus Score (eCRF)";
	
	COMPUTE TEST_PRURSCOR ;
	IF TEST_PRURSCOR = "COHERENT" THEN CALL DEFINE (_COL_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTGREEN]");
	ELSE IF TEST_PRURSCOR = "INCOHERENT" THEN CALL DEFINE (_COL_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTRED]");
	ENDCOMP;

RUN;

			/*>>>>>>>>>>>>>>>>>>>		HAMILTON RATING SCALE		<<<<<<<<<<<<<<<<<<<*/

ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "HAMILTON RATING SCALE");
option missing = "";
PROC REPORT DATA = TAB_HAM_1

	nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "HAMILTON RATING SCALE";

	DEFINE TEST_QSORRES / "Hamilton Rating Reconciliation result";
	DEFINE Findings_value1 / "Hamilton Rating (IWRS)";
	DEFINE QSORRES / "Hamilton Rating (eCRF)";
	
	COMPUTE TEST_QSORRES ;
	IF TEST_QSORRES = "COHERENT" THEN CALL DEFINE (_COL_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTGREEN]");
	ELSE IF TEST_QSORRES = "INCOHERENT" THEN CALL DEFINE (_COL_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTRED]");
	ENDCOMP;

RUN;

			/*>>>>>>>>>>>>>>>>>>>>>>>		FATIGUE SEVERITY SCALE		<<<<<<<<<<<<<<<<<<<<<<*/

ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "FATIGUE SEVERITY SCALE");
option missing = "";
PROC REPORT DATA = TAB_FSS_1

	nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "FATIGUE SEVERITY SCALE";

	DEFINE TEST_QSFSSSCR / "Fatigue Severity Scale Reconciliation result";
	DEFINE Findings_value1 / "Fatigue Severity Scale (IWRS)";
	DEFINE QSFSSSCR / "Fatigue Severity Scale (eCRF)";
	
	COMPUTE TEST_QSFSSSCR ;
	IF TEST_QSFSSSCR = "COHERENT" THEN CALL DEFINE (_COL_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTGREEN]");
	ELSE IF TEST_QSFSSSCR = "INCOHERENT" THEN CALL DEFINE (_COL_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTRED]");
	ENDCOMP;

RUN;

			/*>>>>>>>>>>>>>>>>>>>>>>>>		NUMBER OF FLUSHES		<<<<<<<<<<<<<<<<<<<<<<*/

ODS EXCEL OPTIONS(FROZEN_HEADERS = "3"  EMBEDDED_TITLES = 'YES' SHEET_NAME = "NUMBER OF FLUSHES");
option missing = "";
PROC REPORT DATA = TAB_FLUSHES_1

	nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	title j=c "NUMBER OF FLUSHES";

	DEFINE TEST_QSFLUSHORRES / "Number of Flushes Reconciliation result";
	DEFINE Findings_value1 / "Number of Flushes (IWRS)";
	DEFINE QSFLUSHORRES / "Number of Flushes (eCRF)";
	
	COMPUTE TEST_QSFLUSHORRES ;
	IF TEST_QSFLUSHORRES = "COHERENT" THEN CALL DEFINE (_COL_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTGREEN]");
	ELSE IF TEST_QSFLUSHORRES = "INCOHERENT" THEN CALL DEFINE (_COL_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTRED]");
	ENDCOMP;

RUN;

/*Close file*/

ODS EXCEL CLOSE;

			  /*****************************************************************************/
			 /*************				ADD RECAP STUDY					*******************/
			/*****************************************************************************/

/*%INCLUDE "M:\21 Data Management\DADI\AB19001\Recap Study\Script\Recap Study.sas";*/



