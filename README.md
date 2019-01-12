# ms-office-access-utilities
Utitlity classes for MS Access version 2003-2019

This repository provides a utility class bundle written in VBA that will help you automate complex tasks such as importing, exporting and syncing with other databases. Doing file management with ease. Looking up and resolving contacts in outlook and active directory and much more. Using the database as both local and global settings reposiroty. 
Implementing update and sync protocols with other backends and frontends.

The utility and all other helper classes are exported in their own files and are not part of a particular solution.
These are classes that build on top of the functionality and abstractions that come with VBA.

The main interface class Utility exposes all other helper and utility classes. This is done purely out of convinience for the developer. All utility classes are initiated in the interface class constructor and thus available when the document is opened.
This allows the develper to refer them in the debug console and in any other class just by navigating through Util.<class>.
  
The only utility class specific to MS Access is UtilAccess.cls, however it uses other general utility class that have also been included in the ms-office-excel-utilities repository.

There are currently different versions of the general utility classes becuase the the repositories were not related up to now.
A merge of the general utility classes is on the agenda and will most probalby be available as a seperate repository.
