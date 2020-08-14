import xlwings as xw
import numpy as np
import random
import time


"""
1. Define the path for the Excel workbook
2. Use xlwings and the workbook path to create a workbook object
3. Create sheet objects for each tab in the workbook
"""

path = r"C:\Users\tcregar\Desktop\Sorting Algorythems\Sorting_Algos.xlsm"
wb = xw.Book(path)
heap_sort = wb.sheets["Heap Sort"]


#######################################################################################
#                                 HEAP SORT
#######################################################################################

def max_heapify(array, split, i):

    left = 2 * i + 1
    right = 2 * i + 2

    if left < split and array[left] > array[i]:
        largest = left
    else:
        largest = i

    if right < split and array[right] > array[largest]:
        largest = right

    if largest != i:

        heap_sort.range("hs_av").value = array[i]
        heap_sort.range("hs_cfs").value = array[largest]

        i_cell = f'hs_{i}'
        largest_cell = f'hs_{largest}'

        heap_sort.range(i_cell).value = array[largest]
        heap_sort.range(largest_cell).value = array[i]

        cnt = heap_sort.range("hs_cnt").value + 1
        heap_sort.range("hs_cnt").value = cnt
        time.sleep(.075)

        array[i], array[largest] = array[largest], array[i]
        max_heapify(array, split, largest)


def build_max_heap(array):

    split = len(array)

    for i in range(split, -1, -1):
        max_heapify(array, split, i)

    for i in range(split - 1, 0, -1):

        heap_sort.range("hs_av").value = array[0]
        heap_sort.range("hs_cfs").value = array[i]

        i_cell = f'hs_{i}'
        heap_sort.range("hs_0").value = array[i]
        heap_sort.range(i_cell).value = array[0]

        cnt = heap_sort.range("hs_cnt").value + 1
        heap_sort.range("hs_cnt").value = cnt
        time.sleep(.075)

        array[0], array[i] = array[i], array[0]
        max_heapify(array, i, 0)


def build_heap_data():
    data = list(range(0, 62))
    random.shuffle(data)

    heap_sort.range("hs_array").value = data

    heap_sort.range("hs_av").value = 0
    heap_sort.range("hs_cfs").value = 0
    heap_sort.range("hs_cnt").value = 0

    return data


def run_heap():
    data = heap_sort.range("hs_array").value
    build_max_heap(data)
    return data


#####################################################################################
#
#####################################################################################
