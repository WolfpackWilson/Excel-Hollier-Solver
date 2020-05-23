
# Excel Hollier Solver
The Hollier Solver tool is an Excel plugin that takes a from-to table and solves it using Hollier's first method.

|Contents|
|---|
|1. [Introduction](#introduction)|
|&nbsp;&nbsp;&nbsp;1.1 [What are Hollier's Methods?](#what-are-hollier's-methods)|
|&nbsp;&nbsp;&nbsp;1.2 [Hollier Method 1 Algorithm](#hollier-method-1-algorithm)|
|2. [Installation and Use](#installation-and-use)|
|&nbsp;&nbsp;&nbsp;2.1 [How to Use the Solver](#how-to-use-the-solver)|
|&nbsp;&nbsp;&nbsp;2.2 [How to Add the Solver as an Excel Add-In](#how-to-add-the-solver-as-an-excel-add-in)|
|&nbsp;&nbsp;&nbsp;2.3 [Terms of Service](#terms-of-service)|


# Introduction

## What are Hollier's Methods?
Hollier's methods are algorithms used to order machines for minimizing the backtracking of parts. In the context of manufacturing, it is implemented after separating components of an assembly into part families, splitting them into machine groups using rank order clustering, and then generating from-to tables for each group.

## Hollier Method 1 Algorithm
1. Develop a from-to chart based on part routes. 
1. Calculate the "to" and "from" sums for each machine.
1. Assign the machine position based on minimum "from" or "to" summations.
    - If a tie between two sums in the "from" or "to" category exists, the machine with the lowest from-to ratio is chosen.
    - If the minimum "to" and "from" summations are equivalent, the machines are assigned the next and last positions, respectively.
    - If both "to" and "from" sums are equivalent for the same machine, it is skipped, and the next machine with the minimal sum is chosen.
1. Eliminate the assigned machine from the from-to table and repeat the process until all machines have been assigned to a position.


# Installation and Use

## How to Use the Solver

## How to Add the Solver as an Excel Add-In

## Terms of Service
