# SortMultipleColumns Macro
## Underlying Logic
- The code sorts the active sheet using the `.Sort` method.
- It adds two `.SortFields` to define the sorting criteria:
  - Sorting column A in ascending order
  - Sorting column B in ascending order
- The `.SetRange` method sets the range to be sorted, which is A1:C13.
- The `.Header` option is set to `xlYes` to indicate the header row.
- Finally, `.Apply` executes the sort operation.

## Data Flow
The data within the A1:C13 range is being sorted based on the specified criteria. The sort operation modifies the order of the rows, rearranging them according to the information in columns A and B.

## Process Flow
1. The macro starts by selecting the active sheet.
2. It configures the sorting criteria using the `.SortFields.Add` method, specifying the key columns and sorting order.
3. The `.SetRange` method defines the range to be sorted.
4. The sort operation is executed using the `.Apply` method, which applies the defined criteria to the specified range.
