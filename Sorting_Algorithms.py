import xlwings as xw
import numpy as np
import random


path = r"C:\Users\tcregar\Desktop\Sorting Algorythems\SortingAlgorithms.xlsm"
wb = xw.Book(path)
heap_sort = wb.sheets["Heap Sort"]


def swap(array, start, compare):
    loc = f'hs_{start}'
    cng = f'hs_{compare}'
    array[start], array[compare] = array[compare], array[start]
    heap_sort.range(loc).value, heap_sort.range(
        cng).value = heap_sort.range(cng).value, heap_sort.range(loc).value
    heap_sort.range("hs_root").value = array[start]
    heap_sort.range("hs_child").value = array[compare]


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
        swap(array, i, largest)
        max_heapify(array, split, largest)


def get_new_values():
    data = list(range(0, 63))
    random.shuffle(data)
    heap_sort.range("hs_array").value = data
    heap_sort.range("hs_root").value = data[30]
    heap_sort.range("hs_child").value = data[62]
    heap_sort.range("hs_sorted").value = len(data)-1


def build_max_heap():
    array = heap_sort.range("hs_array").value
    split = len(array)

    for i in range(split, -1, -1):
        max_heapify(array, split, i)

    heap_sort.range("hs_root").value = array[0]
    heap_sort.range("hs_child").value = array[62]


def run_heap_sort():
    array = heap_sort.range("hs_array").value
    split = len(array)

    for i in range(split, -1, -1):
        max_heapify(array, split, i)

    for i in range(split - 1, 0, -1):
        swap(array, 0, i)
        max_heapify(array, i, 0)
        heap_sort.range("hs_sorted").value = array[i]
    heap_sort.range("hs_sorted").value = array[0]

    heap_sort.range("hs_root").value = array[0]
    heap_sort.range("hs_child").value = array[1]
