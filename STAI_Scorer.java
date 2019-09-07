/*
STAI Scorer for STAI-AD Qualtrics outputs used in Miwa Lab SNP Project.
Author: Mark A. Crenshaw
*/



import com.sun.media.sound.InvalidFormatException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;


public class STAI_Scorer {


    /*
     ******* SCORES ********
     */

    //Adult Male Percentiles
    static int[][] state_male19_percentile = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,99,98,98,98,97,96,95,94,94,93,92,91,90,88,86,85,82,80,78,76,73,70,66,64,61,58,55,50,46,44,39,36,31,28,25,19,16,14,12,9,8,6,4}, };
    static int[][] trait_male19_percentile = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,99,99,98,98,98,97,96,96,95,94,94,92,90,88,87,85,83,81,78,76,74,71,69,66,63,59,57,52,48,43,38,33,30,27,24,21,15,12,11,7,4,3}, };
    static int[][] state_male40_percentile = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {100,100,100,100,100,100,100,100,100,100,100,100,100,99,99,98,98,98,98,98,98,97,96,96,95,94,94,93,93,92,90,89,88,87,85,83,81,78,76,72,70,67,64,62,58,56,53,48,43,39,35,27,24,22,19,16,14,12,9,6,5}, };
    static int[][] trait_male40_percentile = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,99,99,99,98,97,97,96,94,93,92,90,89,87,86,84,82,81,78,76,73,68,65,62,60,54,49,44,39,34,28,24,21,18,14,11,8,5,3,1}, };
    static int[][] state_male50_percentile = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,99,99,99,99,98,98,97,97,96,96,96,95,94,94,92,91,89,87,85,84,83,81,79,76,74,72,69,66,64,60,55,52,48,45,40,36,33,28,26,21,18,16,11,9,6}, };
    static int[][] trait_male50_percentile = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,99,98,98,98,97,96,96,94,94,93,92,91,90,88,86,84,81,77,74,71,68,63,61,59,55,49,45,39,36,31,27,24,19,15,11,8,6,3}, };

    //Adult Female Percentiles
    static int[][] state_female19_percentile = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {100,100,100,100,100,100,100,100,100,100,100,100,100,100,99,99,99,98,97,95,95,95,95,94,94,93,92,91,91,89,89,87,85,84,82,81,79,77,76,73,71,71,68,62,59,56,52,48,44,41,40,34,30,21,17,13,10,9,6,3,2}, };
    static int[][] trait_female19_percentile = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,99,98,97,96,95,95,95,94,93,93,93,92,92,90,89,86,86,83,80,76,72,69,66,65,61,59,54,50,47,42,35,29,25,22,18,16,12,9,7,3,3,0}, };
    static int[][] state_female40_percentile = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {100,100,100,100,100,100,100,100,100,100,100,100,100,100,99,99,99,98,96,96,96,96,96,94,94,93,93,93,91,89,87,87,87,85,82,81,78,75,74,72,67,67,67,64,58,55,53,50,49,43,39,33,24,22,19,16,16,13,8,5,3}, };
    static int[][] trait_female40_percentile = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,99,99,99,99,99,99,97,96,95,94,93,92,92,92,90,89,87,87,84,82,80,78,78,74,70,65,63,57,53,50,45,44,37,33,27,22,17,14,11,7,5,2,0}, };
    static int[][] state_female50_percentile = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,99,99,99,99,99,99,99,99,99,99,99,99,99,99,99,99,97,97,93,93,93,93,90,87,85,82,80,76,74,72,69,66,61,59,51,47,37,35,32,31,28,24,22,12,8,5}, };
    static int[][] trait_female50_percentile = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,99,99,99,98,97,97,97,97,96,95,93,92,88,84,83,81,76,73,69,66,59,56,51,44,39,34,31,30,27,23,19,14,8,7}, };

    //Adult Male Standard Scores
    static int[][] state_male19_standard = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {93,92,91,90,89,88,87,86,85,84,83,82,81,80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34}, };
    static int[][] trait_male19_standard = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {96,95,93,92,91,90,89,88,87,86,85,84,83,82,81,80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34}, };
    static int[][] state_male40_standard = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {92,91,90,89,88,87,86,85,84,83,82,81,81,80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,43,42,41,40,39,38,37,36,35}, };
    static int[][] trait_male40_standard = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {101,99,98,97,96,95,94,93,92,90,89,88,87,86,85,84,83,81,80,79,78,77,76,75,74,72,71,70,69,68,67,66,65,63,62,61,60,59,58,57,56,54,53,52,51,50,49,48,47,45,44,43,42,41,40,39,38,36,35,34,33}, };
    static int[][] state_male50_standard = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {94,93,92,91,90,89,88,87,86,85,84,83,82,81,80,80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36}, };
    static int[][] trait_male50_standard = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {102,101,100,99,97,96,95,94,93,92,91,90,88,87,86,85,84,83,82,81,80,78,77,76,75,74,73,72,70,69,68,67,66,65,64,63,61,60,59,58,57,56,55,54,52,51,50,49,48,47,46,45,43,42,41,40,39,38,37,35,34}, };

    //Adult Female Standard Scores
    static int[][] state_female19_standard = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {90,89,88,87,86,85,84,84,83,82,81,80,79,78,77,76,75,74,74,73,72,71,70,69,68,67,66,65,64,64,63,62,61,60,59,58,57,56,55,54,53,53,52,51,50,49,48,47,46,45,44,43,43,42,41,40,39,38,37,36,35}, };
    static int[][] trait_female19_standard = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {96,95,94,93,92,91,90,89,88,87,86,84,83,82,81,80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,42,41,40,39,38,37,36,35,34,33}, };
    static int[][] state_female40_standard = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {90,89,88,87,86,85,84,84,83,82,81,80,79,78,77,76,75,74,74,73,72,71,70,69,68,67,66,65,64,64,63,62,61,60,59,58,57,56,55,55,54,53,52,51,50,49,48,47,46,45,45,44,43,42,41,40,39,38,37,36,36}, };
    static int[][] trait_female40_standard = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {98,97,96,95,94,93,92,91,90,89,88,87,85,84,83,82,81,80,79,78,77,76,75,74,73,71,70,69,68,67,66,65,64,63,62,61,60,59,57,56,55,54,53,52,51,50,49,48,47,46,45,44,42,41,40,39,38,37,36,35,34}, };
    static int[][] state_female50_standard = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {105,104,102,101,100,99,98,97,96,95,94,92,91,90,89,88,86,85,84,83,82,81,80,78,77,76,75,74,73,72,71,69,68,67,66,65,63,62,61,60,59,58,57,55,54,53,52,51,50,49,47,46,45,44,43,42,40,39,38,37,36}, };
    static int[][] trait_female50_standard = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {112,111,109,108,107,106,104,103,102,100,99,98,97,95,94,93,91,90,89,88,86,85,84,82,81,80,79,77,76,75,73,72,71,70,68,67,66,64,63,62,61,59,58,57,55,54,53,52,50,49,48,46,45,44,42,41,40,39,37,36,35}, };

    //College Students Percentile
    static int[][] state_male_college_percentile = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,99,99,98,98,98,97,97,97,97,96,95,94,93,92,92,90,88,86,84,83,81,80,78,75,72,70,68,64,62,58,53,49,46,42,36,30,25,22,19,17,12,9,6,4,2,2}, };
    static int[][] trait_male_college_percentile = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {100,100,100,100,100,100,100,100,100,100,100,100,100,100,99,99,99,99,99,99,99,99,99,97,97,96,95,93,92,90,88,87,85,82,81,79,76,74,71,67,60,57,54,52,49,44,38,35,33,28,22,16,12,10,8,6,3,3,1,1,0}, };
    static int[][] state_female_college_percentile = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {100,100,100,100,100,100,100,100,100,100,99,99,99,99,99,98,97,96,95,94,94,93,92,91,90,89,88,86,85,84,82,80,79,78,75,73,71,69,68,66,63,58,55,52,47,45,42,39,35,31,28,24,20,17,15,12,10,8,6,4,3}, };
    static int[][] trait_female_college_percentile = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {100,100,100,100,100,100,100,100,100,100,99,99,99,99,99,98,98,97,97,96,96,95,94,93,92,91,90,89,87,86,85,83,81,79,76,72,69,66,62,59,53,50,46,42,40,36,32,29,25,21,17,14,10,8,5,3,2,1,0,0,0}, };


    //High School Percentile
    static int[][] state_male_high_percentile = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {100,100,100,100,100,100,100,100,100,100,100,100,100,100,99,99,99,99,99,99,98,98,98,97,96,95,94,93,92,92,88,85,81,80,76,73,68,64,61,59,56,54,48,44,41,38,32,30,26,21,18,15,13,12,10,8,7,5,3,2,2}, };
    static int[][] trait_male_high_percentile = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {100,100,100,100,100,100,100,100,99,99,99,99,99,98,98,98,98,98,98,98,97,97,97,96,95,94,92,90,88,87,85,81,78,75,71,68,65,63,60,56,53,47,45,43,41,36,34,30,25,22,19,16,14,11,9,7,6,4,3,2,2}, };
    static int[][] state_female_high_percentile = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {100,100,100,100,100,100,99,99,99,99,98,97,97,97,96,96,94,94,93,93,92,92,90,89,88,87,86,86,84,82,80,78,76,73,72,68,64,60,58,55,51,49,45,44,41,39,37,35,33,29,27,23,20,18,16,12,9,6,4,4,2}, };
    static int[][] trait_female_high_percentile = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {100,100,100,100,100,100,100,100,100,99,99,99,98,98,98,98,98,97,97,97,96,95,94,93,92,91,90,88,85,83,80,78,76,74,72,69,65,61,59,57,53,48,44,40,36,33,31,27,23,20,18,15,10,9,7,6,5,3,2,1,0}, };

    //Military Percentile
    static int[][] state_military_percentile = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {100,100,100,99,99,99,99,99,98,98,97,97,97,96,95,95,94,93,92,91,90,89,87,86,84,82,81,79,77,75,72,69,67,64,61,58,55,51,48,44,41,38,35,32,28,26,23,21,18,16,14,11,10,8,7,5,4,3,2,1,1}, };
    static int[][] trait_military_percentile = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {100,100,100,100,100,100,100,100,100,100,100,100,100,100,99,99,99,99,99,98,98,97,97,97,96,96,95,94,93,92,90,89,87,85,83,80,78,75,72,69,66,63,59,55,51,46,41,38,34,29,25,21,16,12,10,8,5,4,2,1,1}, };

    //College Students Standard Scores
    static int[][] state_male_college_standard = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {93,92,91,90,89,88,87,86,85,84,83,82,81,80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34}, };
    static int[][] trait_male_college_standard = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {95,94,93,92,91,90,89,88,87,86,85,83,82,81,80,79,78,77,76,75,74,73,71,70,69,68,67,66,65,64,63,62,61,59,58,57,56,55,54,53,52,51,50,49,47,46,45,44,43,42,41,40,39,38,37,36,34,33,32,31,30}, };
    static int[][] state_female_college_standard = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {85,84,83,82,81,80,79,79,78,77,76,75,74,74,73,72,71,70,69,69,68,67,66,65,64,64,63,62,61,60,59,59,58,57,56,55,54,54,53,52,51,50,49,49,48,47,46,45,44,44,43,42,41,40,39,38,38,37,36,35,34}, };
    static int[][] trait_female_college_standard = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {89,88,87,86,85,84,83,82,81,80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30}, };

    //High School Standard Scores
    static int[][] state_male_high_standard = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {92,91,90,89,88,86,85,84,83,82,81,80,79,78,77,76,75,74,73,72,70,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30}, };
    static int[][] trait_male_high_standard = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {88,87,86,85,84,83,82,81,80,79,78,77,76,75,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,37,36,35,34,33,32,31}, };
    static int[][] state_female_high_standard = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {81,80,79,78,78,77,76,75,74,74,73,72,71,71,70,69,68,67,67,66,65,64,64,63,62,61,60,60,59,58,57,57,56,55,54,53,53,52,51,50,50,49,48,47,46,45,45,44,43,43,42,41,40,39,39,38,37,36,36,35,34}, };
    static int[][] trait_female_high_standard = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {87,86,85,84,83,82,81,80,79,78,77,76,75,74,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,43,42,41,40,39,38,37,36,35,34,33,32,31,30}, };

    //Military Standard Scores
    static int[][] state_military_standard = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {80,79,78,77,76,75,75,74,73,72,71,70,70,69,68,67,66,66,65,64,63,62,62,61,60,59,58,57,57,56,55,54,53,52,52,51,50,49,48,48,47,46,45,44,44,43,42,41,40,39,39,38,37,36,35,34,34,33,32,31,30}, };
    static int[][] trait_military_standard = { {80,79,78,77,76,75,74,73,72,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,52,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32,31,30,29,28,27,26,25,24,23,22,21,20},
            {95,94,92,91,90,89,88,87,86,85,84,83,82,81,80,79,78,77,76,75,74,73,71,70,69,68,67,66,65,64,63,62,61,60,59,58,57,56,55,54,53,51,50,49,48,47,46,45,44,43,42,41,40,39,38,37,36,35,34,33,32}, };

    // This program will score a maximum of 300 participants' data at once. Increase for more.
    static final int MAX_BATCH_SIZE = 300;

    /*
     ******* CODE *******
     */

    public static void main(String[] args) throws IOException, InvalidFormatException {



      // This file contains the STAI State and Trait data along with the respective BDQ data for each
      // participant.

        // CHANGE THIS FILE PATH TO CHANGE FILE NAME
      String excelFilePath = "TO_SCORE.xlsx";

      FileInputStream inputStream = new FileInputStream(excelFilePath);

      Workbook workbook = new XSSFWorkbook(inputStream);

      Sheet sheet = workbook.getSheetAt(0);

      String[] secondaryCodes = new String[MAX_BATCH_SIZE];
      int[] score_state = new int[MAX_BATCH_SIZE];
      int[] score_trait = new int[MAX_BATCH_SIZE];
      int[] age = new int[MAX_BATCH_SIZE];
      String[] sex = new String[MAX_BATCH_SIZE];
      String[] student = new String[MAX_BATCH_SIZE];
      String[] military = new String[MAX_BATCH_SIZE];

      int rowCount = 0;

      for (Row row : sheet) { // For each Row.
          Cell secCodeCell = row.getCell(0); // Get the Cell at the Index / Column you want.
          secondaryCodes[rowCount] = (getCellValue(secCodeCell).toString());

          Cell stateCell = row.getCell(1);
          score_state[rowCount] = ((int)stateCell.getNumericCellValue());

          Cell traitCell = row.getCell(2); // Get the Cell at the Index / Column you want.
          score_trait[rowCount] = ((int)traitCell.getNumericCellValue());

          Cell ageCell = row.getCell(3); // Get the Cell at the Index / Column you want.
          age[rowCount] = ((int)ageCell.getNumericCellValue());

          Cell sexCell = row.getCell(4);
          sex[rowCount] = (sexCell.getStringCellValue());

          Cell studentCell = row.getCell(5);
          student[rowCount] = (studentCell.getStringCellValue());

          rowCount++;
      }

        workbook.close();
        inputStream.close();

        // These arrays store the calculated standard scores and percentiles
               /* stateSS = new ArrayList<Object>(),
                stateStuSS = new ArrayList<Object>(), stateStuPercentile = new ArrayList<Object>(),
                traitSS = new ArrayList<Object>(), traitPercentile = new ArrayList<Object>(),
                traitStuSS = new ArrayList<Object>(), traitStuPercentile = new ArrayList<Object>();
                */
        int[] stateSS = new int[MAX_BATCH_SIZE], statePercentile = new int[MAX_BATCH_SIZE], stateStuSS = new int[MAX_BATCH_SIZE], stateStuPercentile = new int[MAX_BATCH_SIZE],
                traitSS = new int[MAX_BATCH_SIZE], traitPercentile = new int[MAX_BATCH_SIZE], traitStuSS = new int[MAX_BATCH_SIZE],
                traitStuPercentile = new int[MAX_BATCH_SIZE];
        String[] outSecondaryCodes = new String[MAX_BATCH_SIZE];
        // Calculate soores

        // State student scores
        for (int i = 0; i < 60; i++)
        {
            for (int j = 0; j < rowCount; j++)
            {
                if (sex[j].equals("Female") && student[j].equals("Yes"))
                {
                    if (state_female_college_standard[0][i] == score_state[j])
                        stateStuSS[j] = state_female_college_standard[1][i];
                }
                else if (sex[j].equals("Male") && student[j].equals("Yes"))
                {
                    if (state_male_college_standard[0][i] == score_state[j])
                        stateStuSS[j] = state_male_college_standard[1][i];
                }
            }
        }

        // Trait student scores
        for (int i = 0; i < 60; i++)
        {
            for (int j = 0; j < rowCount; j++)
            {
                if (sex[j].equals("Female") && student[j].equals("Yes"))
                {
                    if (trait_female_college_standard[0][i] == score_trait[j])
                        traitStuSS[j] = trait_female_college_standard[1][i];
                }
                else if (sex[j].equals("Male") && student[j].equals("Yes"))
                {
                    if (trait_male_college_standard[0][i] == score_trait[j])
                        traitStuSS[j] = trait_male_college_standard[1][i];
                }
            }
        }

        // State student percentiles
        for (int i = 0; i < 60; i++)
        {
            for (int j = 0; j < rowCount; j++)
            {
                if (sex[j].equals("Female") && student[j].equals("Yes"))
                {
                    if (state_female_college_percentile[0][i] == score_state[j])
                        stateStuPercentile[j] = state_female_college_percentile[1][i];
                }
                else if (sex[j].equals("Male") && student[j].equals("Yes"))
                {
                    if (state_male_college_percentile[0][i] == score_state[j])
                        stateStuPercentile[j] = state_male_college_percentile[1][i];
                }
            }
        }

        // Trait student percentiles
        for (int i = 0; i < 60; i++)
        {
            for (int j = 0; j < rowCount; j++)
            {
                if (sex[j].equals("Female") && student[j].equals("Yes"))
                {
                    if (trait_female_college_percentile[0][i] == score_trait[j])
                        traitStuPercentile[j] = trait_female_college_percentile[1][i];
                }
                else if (sex[j].equals("Male") && student[j].equals("Yes"))
                {
                    if (trait_male_college_percentile[0][i] == score_trait[j])
                        traitStuPercentile[j] = trait_male_college_percentile[1][i];
                }
            }
        }


        // State adult scores
        for (int i = 0; i < 60; i++)
        {
            for (int j = 0; j < rowCount; j++) {
                if (sex[j].equals("Female"))
                {
                    if (age[j] < 40 && state_female19_standard[0][i] == score_state[j])
                        stateSS[j] = state_female19_standard[1][i];
                    else if (age[j] >= 40 && age[j] < 50  && state_female40_standard[0][i] == score_state[j])
                        stateSS[j] = state_female40_standard[1][i];
                    else if (age[j] > 50 && state_female50_standard[0][i] == score_state[j])
                        stateSS[j] = state_female50_standard[1][i];

                }
                else if (sex[j].equals("Male"))
                {
                    if (age[j] < 40 && state_male19_standard[0][i] == score_state[j])
                        stateSS[j] = state_male19_standard[1][i];
                    else if (age[j] >= 40 && age[j] < 50  && state_male40_standard[0][i] == score_state[j])
                        stateSS[j] = state_male40_standard[1][i];
                    else if (age[j] > 50 && state_male50_standard[0][i] == score_state[j])
                        stateSS[j] = state_male50_standard[1][i];
                }
            }
        }

        // Trait adult scores
        for (int i = 0; i < 60; i++)
        {
            for (int j = 0; j < rowCount; j++) {
                if (sex[j].equals("Female"))
                {
                    if (age[j] < 40 && trait_female19_standard[0][i] == score_trait[j])
                        traitSS[j] = trait_female19_standard[1][i];
                    else if (age[j] >= 40 && age[j] < 50  && trait_female40_standard[0][i] == score_trait[j])
                        traitSS[j] = trait_female40_standard[1][i];
                    else if (age[j] > 50 && trait_female50_standard[0][i] == score_trait[j])
                        traitSS[j] = trait_female50_standard[1][i];

                }
                else if (sex[j].equals("Male"))
                {
                    if (age[j] < 40 && state_male19_standard[0][i] == score_state[j])
                        traitSS[j] = state_male19_standard[1][i];
                    else if (age[j] >= 40 && age[j] < 50  && state_male40_standard[0][i] == score_state[j])
                        traitSS[j] = state_male40_standard[1][i];
                    else if (age[j] > 50 && state_male50_standard[0][i] == score_state[j])
                        traitSS[j] = state_male50_standard[1][i];
                }


            }
        }

        // Trait adult percentiles
        for (int i = 0; i < 60; i++)
        {
            for (int j = 0; j < rowCount; j++) {
                if (sex[j].equals("Female"))
                {
                    if (age[j] < 40 && trait_female19_percentile[0][i] == score_trait[j])
                        traitPercentile[j] = trait_female19_percentile[1][i];
                    else if (age[j] >= 40 && age[j] < 50  && trait_female40_percentile[0][i] == score_trait[j])
                        traitPercentile[j] = trait_female40_percentile[1][i];
                    else if (age[j] > 50 && trait_female50_percentile[0][i] == score_trait[j])
                        traitPercentile[j] = trait_female50_percentile[1][i];

                }
                else if (sex[j].equals("Male"))
                {
                    if (age[j] < 40 && trait_male19_percentile[0][i] == score_trait[j])
                        traitPercentile[j] = trait_male19_percentile[1][i];
                    else if (age[j] >= 40 && age[j] < 50  && trait_male40_percentile[0][i] == score_trait[j])
                        traitPercentile[j] = trait_male40_percentile[1][i];
                    else if (age[j] > 50 && trait_male50_percentile[0][i] == score_trait[j])
                        traitPercentile[j] = trait_male50_percentile[1][i];
                }


            }
        }
        // State adult percentiles
        for (int i = 0; i < 60; i++)
        {
            for (int j = 0; j < rowCount; j++) {
                if (sex[j].equals("Female"))
                {
                    if (age[j] < 40 && state_female19_percentile[0][i] == score_state[j])
                        statePercentile[j] = state_female19_percentile[1][i];
                    else if (age[j] >= 40 && age[j] < 50  && state_female40_percentile[0][i] == score_state[j])
                    statePercentile[j] = state_female40_percentile[1][i];
                    else if (age[j] > 50 && state_female50_percentile[0][i] == score_state[j])
                        statePercentile[j] = state_female50_percentile[1][i];

                }
                else if (sex[j].equals("Male"))
                {
                    if (age[j] < 40 && state_male19_percentile[0][i] == score_state[j])
                        statePercentile[j] = state_male19_percentile[1][i];
                    else if (age[j] >= 40 && age[j] < 50  && state_male40_percentile[0][i] == score_state[j])
                        statePercentile[j] = state_male40_percentile[1][i];
                    else if (age[j] > 50 && state_male50_percentile[0][i] == score_state[j])
                        statePercentile[j] = state_male50_percentile[1][i];
                }


            }
        }

        System.out.println("State score: " + Arrays.toString(stateSS));
        System.out.println("State percentile: " + Arrays.toString(statePercentile));
        System.out.println("State student score: " + Arrays.toString(stateStuSS));
        System.out.println("State student %: " + Arrays.toString(stateStuPercentile));
        System.out.println("Trait score: " + Arrays.toString(traitSS));
        System.out.println("Trait percentile: " + Arrays.toString(traitPercentile));
        System.out.println("Trait student score: " + Arrays.toString(traitStuSS));
        System.out.println("Trait student %:" + Arrays.toString(traitStuPercentile));


      XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sht = wb.createSheet("Scores");

        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[] {"Secondary Code", "STAIAD_State_Adult_SS", "STAIAD_State_Adult_Percentile",
                "STAIAD_State_Student_SS", "STAIAD_State_Student_Percentile", "STAIAD_Trait_Adult_SS",
                "STAIAD_Trait_Adult_Percentile", "STAIAD_Trait_Student_SS", "STAIAD_Trait_Student_Percentile"} );
       for (int i = 0; i < rowCount; i++)
            data.put("" + (i+2) + "", new Object[] { secondaryCodes[i], stateSS[i], statePercentile[i],
                stateStuSS[i], stateStuPercentile[i], traitSS[i], traitPercentile[i], traitStuSS[i], traitStuPercentile[i]} );



        // Iterate over data and write to sheet
        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset) {
            // this creates a new row in the sheet
            Row row = sht.createRow(rownum++);
            Object[] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr) {
                // this line creates a cell in the next column of that row
                Cell cell = row.createCell(cellnum++);
                if (obj instanceof String)
                    cell.setCellValue((String) obj);
                else if (obj instanceof Integer) {
                    cell.setCellValue((Integer) obj);
                }

            }
        }
        try {
            // this writes the workbook
            FileOutputStream out = new FileOutputStream(new File("scored_data.xlsx")); // change output filename
            wb.write(out);
            out.close();
            System.out.println("scored_data.xlsx written successfully on disk.");
        }
        catch (Exception e) {
            e.printStackTrace();
        }

    }

    static Object getCellValue(Cell cell)
    {
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                return cell.getStringCellValue();
            case Cell.CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue();
            case Cell.CELL_TYPE_NUMERIC:
                return (int)cell.getNumericCellValue();
            default:
                return 0;
        }
    }


}


        /*System.out.println("i: " + i);
                    System.out.println("j: " + j);
                    System.out.println("sex.get(j): " + sex[j]);
                    System.out.println("Score: " + state_female19_percentile[0][i] + "Percentile: " + state_female19_percentile[1][i]);
                    System.out.println("state score: " + score_state[j]);
                    System.out.println("state perc hashmap: " + statePercentile);
                    System.out.println("---");*/

        /*
        for (int i =0; i < 60; i++)
        {

            if (sex[0].equals("Female"))
            {
                System.out.println("Percentile: " + state_male19_percentile[0][i] + " State score: " + score_state[i]);
                if (state_male19_percentile[0][i] == score_state[i])
                    System.out.println("Equal");
            }
        }
        */