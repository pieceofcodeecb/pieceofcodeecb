/*Ods excel example using SAS programming in order to format and visualise SAS outputs in Excel : */


	ods excel file="path.xlsx" options(sheet_name='name' FROZEN_HEADERS="1");

		proc report data=&lib_in..table_&p_ech.;
			define variable_bs / "variable_bs" format=nlnum22.0;
			define variable_bs2 / "variable_bs2" format=nlnum22.0;
			define variable_bs3 / "variable_bs3" format=nlnum22.0;
			column variable_bs variable_bs2 variable_bs3;
			compute variable_bs2;
			if variable_bs="CRITERIA" then do;
				call define(_ROW_, "style", "style=[borderbottomstyle=solid borderbottomcolor=black background=lightblue]");
			end;
			endcomp;
		run;

	ods excel options(sheet_name='name2' FROZEN_HEADERS="1");
			proc report data=&lib_in..table2_&p_ech.;
			define variable_bs / "variable_bs" format=nlnum22.0 style={cellwidth=120};
			define variable_bs2 / "variable_bs2" format=nlnum22.0 style={cellwidth=225};
			define variable_bs3 / "variable_bs3" format=nlnum22.0 style={cellwidth=147};
			compute variable_bs1;
				call define(_COL_, "style", "style=[background=CXFFCC66]");
			endcomp;
			compute variable_bs3;
				call define(_COL_, "style", "style=[background=lightblue]");
			endcomp;
		run;
    
	ods excel close; 
