# gas-sheets-projectrolealloc
Create project role allocation pivot table based on projects names, users and their roles.

## Use

The names of tables at the far end of this code.
Even if you don't ready anything else, please read through the examples below.

## Dependencies

Requires the Sheets API (otherwise you get: Reference error: Sheets is not
defined).  To enable the API, refer to 
  https://stackoverflow.com/questions/45625971/referenceerror-sheets-is-not-defined
Note that this API has a default quota of 100 reads/day, after that it 
throws an "Internal Error".

## Example 1

### Input

  Starting with the projects names, users and their roles in sheet /persons/
  
  | Username | Role    | Type      | Project 1 | Project 2 |
  | -------- | ------- | --------- | --------- | --------- |
  | jvonk    | Student | person    | School    | Java      |  
  | svonk    | Student | person    | School    | Reading   |
  | brlevins | Adult   | person    | BoBo      |           |  
  | tiger    | Pet     | cat       | Purr      |           |
  | ownen    | Pet     | cat       | Sleep     | Purr      |

### Filter

  Filter out the people using `Data > Create a Filter > filter on Type == person`

  | Username | Role    | Type      | Project 1 | Project 2 |
  | -------- | ------- | --------- | --------- | --------- |
  | jvonk    | Student | person    | School    | Java      |  
  | svonk    | Student | person    | School    | Reading   |
  | brlevins | Adult   | person    | BoBo      |           |  
  
### Run the script

  Running `OnPivot.main()` generates a sheet called /role-alloc-raw/ and
  the pivot table /role-alloc/.
  
    OnPivot.main(srcColumnLabels = ["Project*", "Username", "Role" ],
                 srcSheetName = "persons",
                 pvtSheetName = "role-alloc");
  
### Intermetiate output

  raw (/role-alloc-raw/)

  | Project  | Username | Role    | Ratio |
  | -------- | -------- | ------- | ----- |
  | School   | jvonk    | Student | 0.5   | 
  | Java     | jvonk    | Student | 0.5   | 
  | School   | svonk    | Student | 0.5   | 
  | Reading  | svonk    | Student | 0.5   |
  | All else | cvonk    | Adult   | 1.0   |
  | BoBo     | brlevins | Adult   | 0.5   |

### Output

  pvt (/role-alloc/)

  | Project | Username | Student | Adult |
  | ------- | -------- | ------- | ----- |
  | School  | jvonk    | 0.5     |       |
  |         | svonk    | 0.5     |       |
  | Java    | jvonk    | 0.5     |       |
  | Reading | svonk    | 0.5     |       |
  | BoBo    | brlevins |         | 1.0   |
  
## Example 2

### Input

  Starting with the projects names, users and their roles in sheet "persons"
  
  | Username | Role    | Type      | Project 1 | Project 2 |
  | -------- | ------- | --------- | --------- | --------- |
  | jvonk    | Student | person    | School    | Java      |  
  | svonk    | Student | person    | School    | Reading   |
  | brlevins | Adult   | person    | BoBo      |           |  

### Themes

  The projects are grouped into themes using the sheet "themes"
  
  | Project  | Theme  |
  | -------- | ------ |
  | School   | Study  |
  | Java     | Study  |
  | Laundry  | Chores | 
  | All else | Chores |
  | Chef     | Chores |
  | Reading  | Relax  |
  | BoBo     | Work   |

### Run the script

  Running `OnPivot.main()` generates a sheet called /role-alloc-raw/ and
  the pivot table /role-alloc/.
  
    OnPivot.main(srcColumnLabels = ["Project*", "Username", "Role" ],
                 srcSheetName = "persons",
                 pvtSheetName = "role-alloc",
                 theSheetName = "themes");

### Intermetiate output

  raw (/role-alloc-raw/)

  | Theme  | Project  | Username | Role    | Ratio |
  | ------ | -------- | -------- | ------- | ----- |
  | Study  | School   | jvonk    | Student | 0.5   | 
  | Study  | Java     | jvonk    | Student | 0.5   | 
  | Study  | School   | svonk    | Student | 0.5   | 
  | Relax  | Reading  | svonk    | Student | 0.5   |
  | Chores | All else | cvonk    | Adult   | 1.0   |
  | Work   | BoBo     | brlevins | Adult   | 0.5   |
  | Chores | Chef     | brlevins | Adult   | 0.5   |

### Output

  pvt (/role-alloc/)

  | Theme   | Project | Username | Student | Adult |
  | ------- | ------- | -------- | ------- | ----- |
  | Study   | School  | jvonk    | 0.5     |       |
  |         |         | svonk    | 0.5     |       |
  |         | Java    | jvonk    | 0.5     |       |
  | Relax   | Reading | svonk    | 0.5     |       |
  | Work    | BoBo    | brlevins |         | 1.0   |
