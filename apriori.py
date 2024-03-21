# -*- coding: utf-8 -*-
import openpyxl
from collections import defaultdict

def ReadData(File):
    workbook = openpyxl.load_workbook(File)
    sheet = workbook.active
    data = []
    for row in sheet.iter_rows(values_only=True):
        transaction = [str(item) for item in row if item is not None]
        data.append(transaction)
    return data
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


def GenerateCandidates(prev_itemsets, k):
    candidates = set()
    for itemset1 in prev_itemsets:
        for itemset2 in prev_itemsets:
            if len(itemset1.union(itemset2)) == k:
                candidates.add(itemset1.union(itemset2))
    return candidates
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


def CollectFrequent_itemsets(data, min_support):
    item_counts = defaultdict(int)
    for transaction in data:
        for item in transaction:
            item_counts[frozenset([item])] += 1

    TransactionsCount = len(data)
    min_support_count = min_support * TransactionsCount
    frequent_itemsets = []
    k = 1
    while True:
        candidates = GenerateCandidates(set(item_counts.keys()), k + 1)
        frequent_itemsets_k = set()
        for candidate in candidates:
            count = 0
            for transaction in data:
                if candidate.issubset(transaction):
                    count += 1
            if count >= min_support_count:
                frequent_itemsets_k.add(candidate)
        if not frequent_itemsets_k:
            break
        frequent_itemsets.extend(frequent_itemsets_k)
        k += 1

    return frequent_itemsets

#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
File = "Data.xlsx" 
min_support = 0.5 

data = ReadData(File)
frequent_itemsets = CollectFrequent_itemsets(data, min_support)

print("Frequent Itemsets:")
for itemset in frequent_itemsets:
    print(itemset)

