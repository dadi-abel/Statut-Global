/*******************************************************************
PURPOSE       AB15003 efficacy parameters tracker      

AUTHOR        Dadi Abel DIEDHIOU
	
VERSION       n°3
	
DATE          20210715
/*******************************************************************/

libname Metrics "H:\AB15003\01_Data Bases\Efficacy_parameters";
option fmtsearch =(Metrics);

libname AB15003 "H:\AB15003\01_Data Bases\01_LS_Prod";
option fmtsearch =(AB15003);

PROC CONTENTS DATA=Metrics._ALL_ OUT=LISTE_TABLE /*NOPRINT*/;
RUN;


data REP_NORMALIZED_DATA;
format vd date9.;
set metrics.REP_NORMALIZED_DATA;
dov = compress(dov,'-');
VD = input(dov,date9.);
run;

proc sort data=REP_NORMALIZED_DATA out=abc(rename=('HAMD score'n=HAMD))/*(drop=dov)*/;
by subjid vd;
run;

/********** DETECTED SCREEN FAIL **********/
proc sort data=Metrics.screen_fail Out=Screen_Failed; by subjid; run;

data final_scr_fail;
merge abc Screen_Failed;
by subjid;
if scr_fail = "Yes" then delete;
drop Site Patient_number VSDAT1 Diff SCR_FAIL;
run;

/********** DISC YN **********/
data Rep_studydisc;
set AB15003.Rep_studydisc; 
if DSTERM ne "" then DISC="Yes";
run;

proc sort data=Rep_studydisc; by subjid; run;

proc sql ;
create table final_scr_F_disc as
select A.*,
Case 
when B.DISC="Yes" then B.DISC
when B.DISC="" then "No"
end as DISC_YN
from final_scr_fail A 
left join Rep_studydisc B
on A.subjid=B.subjid ;
run;

/********** Last Ordered Visit performed + Final Visit exist YN **********/
data Rep_vd;
set AB15003.Rep_vd;
dov = compress(vsdat,'-');
VD = input(dov,date9.);
format vd date9.;
run;

proc sort data=Rep_vd; by subjid VD; run;

data Rep_vd_FV;
set Rep_vd;
where SUBJEVENTNAME="Final Visit";
run;

data LV_ordered;
set rep_vd;
where SUBJEVENTNAME^="Final Visit" and VISPER^="No";
if last.subjid;
by subjid;
run;

/********** Last Intake Date **********/
data admin_end_date(keep=subjid Last_intake RECORDNUM ecdose);
set AB15003.rep_studymed_admin;
L_I = compress(ECENDAT,'-');
Last_intake = input(L_I,date9.);
format Last_intake date9.;
run;

proc sort data=admin_end_date; by subjid RECORDNUM Last_intake; run;

data last_admin_date;
set admin_end_date;
if Last_intake = . then ongoing_intake="Yes";
if last.subjid;
by subjid;
run;

/********** Rando Date **********/

data rando_date(keep=subjid RANDYN Rando_Dat);
set AB15003.rep_randomization;
Rando = compress(DSSTDAT,'-');
Rando_Dat = input(Rando,date9.);
format Rando_Dat date9.;
run;

proc sort data=rando_date; by subjid Rando_Dat; run;


/*merge*/
proc sql ;
create table final_LV_FV as
select A.*, B.SUBJEVENTNAME as Last_Ord_Visit, 
Case 
when C.VISPER="Yes" then C.VISPER
when C.VISPER="" then "No"
end as FV_YN,
C.VD as FV_Date,
Last_intake, ongoing_intake,
Rando_Dat,
Case 
when Last_intake is not null and Rando_Dat is not null then put(datdif(Rando_Dat,Last_intake,'ACT/ACT'),4.) 
end as L_I_Post_Rando,
Case 
when FV_Date is not null and Rando_Dat is not null then put(datdif(Rando_Dat,FV_Date,'ACT/ACT'),4.) 
end as FV_Post_Rando,
Case 
when FV_Date is not null and Last_intake is not null then put(datdif(Last_intake,FV_Date,'ACT/ACT'),4.) 
end as FV_Post_Last_Int
from final_scr_F_disc A 
left join LV_ordered B
on A.subjid=B.subjid
left join Rep_vd_FV C
on A.subjid=C.subjid
left join last_admin_date D
on A.subjid=D.subjid
left join rando_date E
on A.subjid=E.subjid;
run;



proc sort data=final_LV_FV; by subjid vd SUBJEVENTNAME; run;

proc transpose data=final_LV_FV out=xyz(drop=_label_);
by studyid subjid DISC_YN Last_Ord_Visit FV_YN L_I_Post_Rando FV_Post_Rando FV_Post_Last_Int siteid sitename vd SUBJEVENTNAME;
var  dov HAMD Flushes Pruritus fss fis;
quit;

proc sort data=xyz;
by studyid subjid DISC_YN Last_Ord_Visit FV_YN L_I_Post_Rando FV_Post_Rando FV_Post_Last_Int siteid _name_;
run;

proc transpose data=xyz out=bcv;
by studyid subjid DISC_YN Last_Ord_Visit FV_YN L_I_Post_Rando FV_Post_Rando FV_Post_Last_Int siteid sitename _name_;
var SUBJEVENTNAME col1;
id SUBJEVENTNAME;
idlabel SUBJEVENTNAME;
run; 

data final (drop=_label_ studyid siteid  rename=(_name_=Handicaps /*'Baseline Visit'n=Baseline 'Week 4'n=W4 'Week 8'n=W8 'Week 12'n=W12 'Week 16'n=W16 'Week 20'n=W20*/ 'Final Visit'n=Final_Visit));
set bcv;
if ^missing(_label_) then delete;
run;


/*Move the column 'final visit' to the end without knowing the rest of columns*/ 
/* delete non confirmed handicaps at baseline*/
data final1;
set final(drop= Final_Visit);
set final(keep= Final_Visit);
if Handicaps="Pruritus" /*and Screening<9*/ and (missing(Baseline) or Baseline<9) then delete;
if Handicaps="Flushes" /*and Screening<8*/ and (missing(Baseline) or Baseline<8) then delete;
if Handicaps="HAMD" /*and Screening<19*/ and (missing(Baseline) or Baseline<19) then delete;
if Handicaps="FSS" and Baseline<36 then delete;
if Handicaps="FIS" and Baseline<75 then delete;
run;

PROC CONTENTS DATA=final1 OUT=contents /*NOPRINT*/;
RUN;
proc sql /*noprint*/;
 select name into :varlist separated by ' ' 
 from contents 
 order by varnum
 ;
 select count(*) into: loop from contents;
quit;
%put &varlist.;


data final2 (drop=i);
/*format vd date9.;*/
set final1;
array variables {*} &varlist. ;
If Handicaps^='DOV' then
do;
W4_Evolution_Rate = put((W004-Baseline)/Baseline, percentn7.0) ;
W8_Evolution_Rate = put((W008-Baseline)/Baseline, percentn7.0) ;
W12_Evolution_Rate = put((W012-Baseline)/Baseline, percentn7.0) ;
W16_Evolution_Rate = put((W016-Baseline)/Baseline, percentn7.0) ;
W20_Evolution_Rate = put((W020-Baseline)/Baseline, percentn7.0) ;
W24_Evolution_Rate = put((W024-Baseline)/Baseline, percentn7.0) ;
FinalV_Evolution_Rate = put((Final_Visit-Baseline)/Baseline, percentn7.0) ;
end;
else 
	Do i=1 to dim(variables);
		variables{i}=' ';
	end;
run;

data final3 (drop=W4_Evolution_Rate W8_Evolution_Rate W12_Evolution_Rate W16_Evolution_Rate W20_Evolution_Rate W24_Evolution_Rate FinalV_Evolution_Rate);
set final2;
label Handicaps="Handicap";
run;



/*Export to Excel (way to overwrite existing excel file)*/
/*
PROC EXPORT DATA = final3
         OUTFILE = "H:\AB15003\01_Data Bases\Efficacy_parameters\Efficacy_parameters_%sysfunc(today(),yymmddn.).xlsx"
            DBMS = EXCEL
             REPLACE ;
     SHEET = "Efficacy" ;
RUN ; */

ODS excel  
	FILE = "H:\AB15003\01_Data Bases\Efficacy_parameters\Efficacy_parameters_%sysfunc(date(), date9.).xlsx"
	options( embedded_titles="yes" sheet_name="Efficacy" orientation='landscape');
	Footnote j=l "Output on %sysfunc(date(), date9.)";

/*>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<*/
ODS excel OPTIONS(FROZEN_HEADERS = "1" FROZEN_ROWHEADERS = "9" autofilter = "All" sheet_interval='none' SHEET_NAME = 'Efficacy');
PROC REPORT DATA = final3

	nowd headline headskip missing ls=256 ps=90 
	Style(header)=[  cellspacing = 3  borderwidth = 1 bordercolordark = black background=grey foreground=white just=center vjust=center font_size=2 font_weight=bold]
	Style(column)=[CELLWIDTH=150  cellspacing = 3  borderwidth = 1 bordercolordark = black background=white just=center vjust=center font_size=2];
	/*title j=c "SEX";*/

	DEFINE SUBJID / "Subject ID";
	DEFINE DISC_YN / "Discontinuation YN";
	DEFINE Last_Ord_Visit / "Last Ordered Visit performed";
	DEFINE FV_YN / "Final Visit YN";
	DEFINE L_I_Post_Rando / "Last Intake days Post Rando";
	DEFINE FV_Post_Rando / "FV days Post Rando";
	DEFINE FV_Post_Last_Int / "FV days Post Last Intake";
	
	COMPUTE SUBJID ;
	IF SUBJID = "" THEN CALL DEFINE (_ROW_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTGREY]");
	/*ELSE IF TEST_SEX = "INCOHERENT" THEN CALL DEFINE (_COL_, "STYLE", "STYLE = [BACKGROUNDCOLOR = LIGHTRED]");*/
	ENDCOMP;

RUN;
ODS excel CLOSE;


/*second way to export to excel*/

/*libname xls excel "H:\AB15003\01_Data Bases\Metrics\Efficacy_parameters.xlsx";*/
/**/
/*data xls.Efficacy;*/
/*set final2;*/
/*run;*/
/**/
/*libname xls clear;*/
