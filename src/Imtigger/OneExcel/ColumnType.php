<?php
namespace Imtigger\OneExcel;

class ColumnType
{
    const GENERAL = 0;
    const STRING = 1;
    const NUMERIC = 10;
    const INTEGER = 11;
    const BOOLEAN = 20;
    const DATE = 40;
    const TIME = 41;
    const DATETIME = 42;
    const FORMULA = 50;
    const NULL = 99;
}