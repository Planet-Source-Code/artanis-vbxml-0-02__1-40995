vbXML 0.01 - 0.02 Readme
- - - - - - - - - - - - -

1) Whats New in 0.02
2) Complete Features List
3) Future Plans

- - - - - - - - - - - - -
1) Whats New in 0.02

In the latest release of vbXML I havent added very much.  There are about 3 (three) new functions/methods for use:
	- GetChildValue (strQuery As String, Index As Long)
		- Returns the value of a child node of the parent specified by strQuery
		- Useful for XML files where the content is unknown (like in the BlackBook example I have included)
	- GetChildAttribute (strQuery As String, strName As String, Index As Long)
		- Returns the value of the attribute of the child note at point Index.
		- Attribute name is specified by strName
	- GetChildName (strQuery As String, Index As Integer)
		- Returns the name of the child node at Index
		- Useful for entering/retreving information to/from an XML file where the node names are unknown (with the 		  exception of the very first node)

These functions (along with the SetColorValue and GetColorValue functions) are the functions that surpass the power of the CGoXML class wrapper.  vbXML is also a bit more stable (now that 0.02 has been released, which is a very stable build).
- - - - - - - - - - - - -
2) Complete Features List

	- 0.01
		- Created vbXML
		- Includes basic functions/methods for reading data from a file and writing data to a file.
		- Some more advanced functions/methods such as GetColorValue and SetColorValue
		- Not very flexible
			- No support for reading unknown child nodes
			- Simple structure
		- Great front-end for the MSXML Library
			- Took away most (with the exception of handling child nodes) of the work of MSXML
	- 0.02
		- Includes all functions/methods from the 0.01 release
		- Added more advanced functions
			- GetChildValue
			- GetChildAttribute
			- GetChildName
		- Advanced functions handle child nodes for the developer
		- More flexibility (with regards to child nodes)
			- Still fairly simple structure (but simple is better)
		- Fixed the NodeCount function
			- Returns the correct node count now
- - - - - - - - - - - - -
3) Future Plans

	- 0.03
		- Add support for creating easy skin files for applications
			- Maybe only begin with a few functions
				- GetColorValue and SetColorValue are the prefiguration of skin management
		- Minor (or major) bug fixes
			- Only if needed
- - - - - - - - - - - - -