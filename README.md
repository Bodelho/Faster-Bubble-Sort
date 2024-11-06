# Faster Bubble Sort

This is an alternate implementation of the classic "Bubble Sort" algorithm. This alternate algorithm works pretty much the same way the classic "Bubble Sort" algorithm does. It aims at improving the classic "Bubble Sort" algorithm performance; therefore from now on it is dubbed "Smart Bubble Sort".

The classic "Bubble Sort" algorithm repeatedly goes FORWARD over the to-be-sorted array from start to finish, swapping adjacent array elements when the values in those adjacent array elements are "out of order". The sort task is over when the to-be-sorted array is scanned from start to finish and no array elements are swapped.

The "Smart Bubble Sort" algorithm also goes FORWARD over the to-be-sorted array from start to finish, but just once. When the values in back-to-back array elements are found to be "out of order", the array elements are swapped in the array; then the algorithm stops moving FORWARD over the array and it starts moving BACKWARDS, comparing and swapping array elements as required until the "swapped down" value is fit into its right ordering place among the preceding already sorted values. As soon as the "swapped down" value is fit into its right ordering place, the algorithm stops moving BACKWARDS over the array and restarts the FORWARD move at the last "swapped up" array element. The sort task is over when the last array element is reached.

## How does it work?

Be the to-be-sorted array "Dim Data(0 To N)" where N = 6. Be "I" the subscript for an array element "Data(I)" that is being compared to its next neighbor element "Data(J)" (where "J = I + 1"). Be "D" the array subscript for a "swapped down" array element ("0 <= D <= N - 1"), and be "U" the array subscript for a "swapped up" array element ("1 <= U <= N"). Be array Data() unsorted values
```
 +---+---+---+---+---+---+---+
 | 1 | 3 | 5 | 2 | 6 | 7 | 4 |   unsorted data to be ascendlingly sorted
 +---+---+---+---+---+---+---+
   0   1   2   3   4   5   6
```
The algorithm goes forward from subscript "0" to subscript "2" and no array elements swapping takes place because values "1", "3", and "5" are already in an ascending order. When "Data(2)" is compared to "Data(3)", those elements' values are swapped because "Data(2) > Data(3)". At this point: "D = 2", "U = 3", and 3 comparisons have been made.
```
 +---+---+---+---+---+---+---+
 | 1 | 3 | 2 | 5 | 6 | 7 | 4 |
 +---+---+---+---+---+---+---+  before: I = 0; J = 1
   0   1   2   3   4   5   6    after:  I = 2; J = 3; D = 2; U = 3
```
Now the algorithm starts moving backwards comparing "Data(2)" to "Data(1)", and have them swapped because "Data(1) > Data(2)".
```
 +---+---+---+---+---+---+---+
 | 1 | 2 | 3 | 5 | 6 | 7 | 4 |
 +---+---+---+---+---+---+---+  before: I = 1; J = 2
   0   1   2   3   4   5   6    after:  D = 1; U = 3
```
Next, "Data(1)" is compared to "Data(0)" and NO swap occurs because "Data(1) > Data(0)", meaning that the "swapped down" value ("2") has been moved into its right spot among already sorted data (using 2 comparisons).
```
 +---+---+---+---+---+---+---+
 | 1 | 2 | 3 | 5 | 6 | 7 | 4 |
 +---+---+---+---+---+---+---+  before: I = 0; J = 1
   0   1   2   3   4   5   6    after:  D = 1; U = 3
```
Then the algorithm restarts moving forward at the "swapped up" array element "Data(U)" (where "U = 3"). No array elements swapping occurs because values "5", "6", and "7" are already in ascending order. When "Data(5)" is compared to "Data(6)", those elements' values are swapped because "Data(5) > Data(6)". At this point: "D = 5", "U = 6", and 3 additional comparisons have been made.
```
 +---+---+---+---+---+---+---+
 | 1 | 2 | 3 | 5 | 6 | 4 | 7 |
 +---+---+---+---+---+---+---+  before: I = 3; J = I + 1 = 4
   0   1   2   3   4   5   6    after:  I = 5; J = 6; D = 5; U = 6
```
Now the algorithm starts moving backwards comparing "Data(4)" to "Data(5)", and have them swapped because "Data(4) > Data(5)".
```
 +---+---+---+---+---+---+---+
 | 1 | 2 | 3 | 5 | 4 | 6 | 7 |
 +---+---+---+---+---+---+---+  before: I = 4; J = 5
   0   1   2   3   4   5   6    after:  D = 4; U = 6
```
"Data(3)" is compared to "Data(4)", and the array elements are swapped because "Data(3) > Data(4)".
```
 +---+---+---+---+---+---+---+
 | 1 | 2 | 3 | 4 | 5 | 6 | 7 |
 +---+---+---+---+---+---+---+  before: I = 3; J = 4
   0   1   2   3   4   5   6    after:  D = 3; U = 6
```
Next, "Data(2)" is compared to "Data(3)" and NO swap occurs because "Data(3) > Data(2)", meaning that the "swapped down" value ("4") has been moved into its right spot among already sorted data (using 3 comparisons).
```
 +---+---+---+---+---+---+---+
 | 1 | 2 | 3 | 4 | 5 | 6 | 7 |
 +---+---+---+---+---+---+---+  before: I = 2; J = 3
   0   1   2   3   4   5   6    after:  D = 3; U = 6
```
Then the algorithm restarts moving forward at the "swapped up" array element "Data(U)" (where "U = 6"). "Data(6)" is the last array element, and the sort task is finished using 11 comparisons ("3 forward + 2 backwards + 3 forward + 3 backwards = 11").

## Classic "Bubble Sort"

If the classic "Bubble Sort" algorithm were used instead to sort the same data set, there would be 18 comparisons to have the data sorted:
```
 +---+---+---+---+---+---+---+
 | 1 | 3 | 5 | 2 | 6 | 7 | 4 |   unsorted data to be ascendlingly sorted
 +---+---+---+---+---+---+---+
   0   1   2   3   4   5   6
```
```
 +---+---+---+---+---+---+---+
 | 1 | 3 | 2 | 5 | 6 | 4 | 7 |   I = 0; J = 0 To (N - I - 1 = 5)
 +---+---+---+---+---+---+---+   swapped 2x3, 5x6: 6 comparisons
   0   1   2   3   4   5   6
```
```
 +---+---+---+---+---+---+---+
 | 1 | 2 | 3 | 5 | 4 | 6 | 7 |   I = 1; J = 0 To (N - I - 1 = 4)
 +---+---+---+---+---+---+---+   swapped 4x5: 5 comparisons
   0   1   2   3   4   5   6
```
```
 +---+---+---+---+---+---+---+
 | 1 | 2 | 3 | 4 | 5 | 6 | 7 |   I = 2; J = 0 To (N - I - 1 = 3)
 +---+---+---+---+---+---+---+   swapped 3x4: 4 comparisons
   0   1   2   3   4   5   6
```
```
 +---+---+---+---+---+---+---+
 | 1 | 2 | 3 | 4 | 5 | 6 | 7 |   I = 3; J = 0 To (N - I - 1 = 2)
 +---+---+---+---+---+---+---+   NO SWAP / SORTED: 3 comparisons
   0   1   2   3   4   5   6
```

## Bottom line

According to some benchmarking carried out on both algorithms using randomly ordered to-be-sorted data sets, "Smart Bubble Sort" performs twice faster than classic "Bubble Sort". Performance of both is the same in a "best case scenario" (that is, already sorted data) and in a "worst case scenario" (that is, reversely ordered data).




