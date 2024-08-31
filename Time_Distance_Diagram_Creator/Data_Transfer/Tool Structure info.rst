=========================
**PSET for PowerFactory**
=========================

The PSET for PowerFactory tool consists of the following four folders:

- calc_sw_interface (collection of the functions interfacing the PSET logic
  with DIgSILENT PowerFactory).
- pset (logic core of the relay setting evaluation tool).
- psetdigsilent_test (the suite of test scripts based on Pytest which are used
  to test the PSET functionalities).
- ui (the PSET graphical user interface python files based on the WX library).


*calc_sw_interface*
-------------------

It has been conceived as a collection of functions which allow to get and set
the power system element parameters, run LDF/SHC calculations, and retrieve
information regarding the connection between the power system elements. 


*pset*
------

The folder contains the following 6 python files:

- assessment.py (code which evaluates the relay trip times calculated by
  pset_logic.py and applies the setting validation rules).

- branch.py (definition of the Branch class which represent a list of
  connected elements modelling a feeder between two bus bars).

- EPRI_PSET_PowerFactory.py (it contains the PSET entry function).

- grid.py (file containing  the Grid class which implements the power system
  layout recognizing logic procedures).

- pset_logic (it contains the PSET main routine which calls in turn all other
  components to create the system layout, run the required SHC, collect the
  tripping times, apply the setting validation rules and output the results).

- report_maker.py (it contains the Report_maker class which implements the
  functions to create the XML file containing a summarized report of the
  results obtained applying the validation rules).


*psetdigsilent_test*
--------------------

Please refer to the ''Test suite description.txt'' file saved inside the
''psetdigsilent_test'' folder.


*ui*
----

The folder contains the following 6 python files:

- about_dlg.py (it contains the implementation of the class representing the
  ''About'' dialog which is displayed selecting the ''Help | About'' menu item).

- busbar_selector_dlg.py (it contains the ''BusbarSelector'' class which
  implements the dialog which allows to select one of the busbars available in
  the power system. The dialog is displayed when the user pushes the ''Select
  Busbar to Study'' button in the PSET main dialog).

- main_panel.py (complete definition of the graphical items present in the
  PSET main dialog, relevant data ranges, and dialog event driven functions).

- main_tab (tabular graphical item hosting the main_panel. No other function
  defined there).

- main_window (basic infrastructure for the PSET main dialog. It manages the
  menu event driven functions, and the ''onFocus'' and ''onExit'' function).
 
 -splash_screen (temporary disclaimer screen showed only for the pre release
 version).



 