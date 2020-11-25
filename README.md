# Visualising sorting algorithms in VBA

After recently changing my [BetterArray](https://github.com/Senipah/VBA-Better-Array) project to use [Timsort](https://en.wikipedia.org/wiki/Timsort) as the default sorting algorithm, I thought it would be fun to quickly throw together a file to visualise the difference between Timsort and the previous default [Quicksort](https://en.wikipedia.org/wiki/Quicksort).

I also threw in a quick [Bubble sort](https://en.wikipedia.org/wiki/Bubble_sort) implementation to highlight just how inefficient it is in comparison.

Here's a video of them in action.

Timsort is slower in the video than Quicksort but in my testing profiling Timsort was about the same if not marginally faster than Quicksort, depending on what is being sorted, so the difference in the video is likely related to how the chart is updated. Timsort is also a stable sort, which was Quicksort is not.

The difference in speed between the Quicksort implementations and Bubble sort is roughly accurate though; they both use the same Swap method to update the chart. So yes, Bubble Sort really is that slow!

Here's a link to the Excel file this was done in if anyone wants to run it themselves.

