# OpenOfficeJSONExporter 

A macro script writen in JavaScript that export spreadsheets into JSON format!

There are two scripts:
- ToArrayOfObject
- ToObjectOfObject

////////////////////////////////////////
Example:
___________________________
|__id__|__name__|_surname_|
|_a123_|_Plus___|_Pingya__|
|_b456_|_Tony___|_Stark___|

////////////////////////////////////////
Run ToArrayOfObject will get:

[
  {
    "id": "a123",
    "name": "Plus",
    "surname": "Pingya"
  },
  {
    "id": "b456",
    "name": "Tony",
    "surname": "Stark"
  }
]

///////////////////////////////////////
Run ToOjectOfObject will get:
{
  "a123": {
    "name": "Plus",
    "surname": "Pingya"
  },
  "b456": {
    "name": "Tony",
    "surname": "Stark"
  }
}

Good Luck!
