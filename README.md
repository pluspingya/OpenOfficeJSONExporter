# OpenOfficeJSONExporter 

A macro script written in JavaScript that export spreadsheets into JSON format!

There are two scripts:
- ToArrayOfObject
- ToObjectOfObject

### Example:
<pre>
___________________________
|__id__|__name__|_surname_|
|_a123_|_Plus___|_Pingya__|
|_b456_|_Tony___|_Stark___|
</pre>

Run ToArrayOfObject will get:
<pre>
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
</prev>

Run ToObjectOfObject will get:
<pre>
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
</pre>
Good Luck!
