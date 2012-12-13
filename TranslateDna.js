/*
JavaScript object that translates a DNA input sequence into a protein sequence.
It uses the "Combination Constructor/Prototype Pattern" as described in the book "Professional JavaScript For Web Developers"
by Nicholas Zakas.
The code is described in link  "http://www.javascript-spreadsheet-programming.com/2012/12/object-oriented-javascript-example.html".
The scientific background is given in "http://www.javascript-spreadsheet-programming.com/2012/11/using-github-for-javascript-and-vba.html"
USE:
Useful if you wish to translate a DNA sequence that contains IUPAC ambiguity codes.
To use (Executed by Node.js): Append the three code lines below to the code file and execute from the command line with:
node TranslateDna.js

var seq = 'CCTKAGATCACTCTTTGGCAACGACCCCTCGTCACAATAAAGATAGGGGGGCAACTAAAGGAAGCTCTATTAGATACAGGAGCAGATGATACAGTATTAGAAGAAATGAATTTGCCAGGAAGATGGAAACCAAAAATGATAGGGGGAATTGGAGGTTTTATCAAAGTAAGACAGTATGATCAGATACTCATAGAAATCTGTGGACATAAAGCTATAGGTACAGTATTAATAGGACCTACACCTGTCAACATAATTGGAAGAAATCTGTTGACTCAGCTTGGTTGCACTTTAAATTTT';
var trans = new TranslateDna(seq);
console.log(trans.getAminoAcids());

Passes JSLint without error when using the default settings.
*/

// Constructor function sets instance variables.
function TranslateDna(dnaSeq) {
    'use strict';
    this.dnaSeq = dnaSeq.toUpperCase();
    this.UNKNOWN = 'X';
    this.translateTable = {'GCT': 'A', 'GCC': 'A', 'GCA': 'A', 'GCG': 'A', 'CGT': 'R', 'CGC': 'R', 'CGA': 'R', 'CGG': 'R', 'AGA': 'R', 'AGG': 'R', 'AAT': 'N', 'AAC': 'N', 'GAT': 'D', 'GAC': 'D', 'TGT': 'C', 'TGC': 'C', 'CAA': 'Q', 'CAG': 'Q', 'GAA': 'E', 'GAG': 'E', 'GGT': 'G', 'GGC': 'G', 'GGA': 'G', 'GGG': 'G', 'CAT': 'H', 'CAC': 'H', 'ATT': 'I', 'ATC': 'I', 'ATA': 'I', 'TTA': 'L', 'TTG': 'L', 'CTT': 'L', 'CTC': 'L', 'CTA': 'L', 'CTG': 'L', 'AAA': 'K', 'AAG': 'K', 'ATG': 'M', 'TTT': 'F', 'TTC': 'F', 'CCT': 'P', 'CCC': 'P', 'CCA': 'P', 'CCG': 'P', 'TCT': 'S', 'TCC': 'S', 'TCA': 'S', 'TCG': 'S', 'AGT': 'S', 'AGC': 'S', 'ACT': 'T', 'ACC': 'T', 'ACA': 'T', 'ACG': 'T', 'TGG': 'W', 'TAT': 'Y', 'TAC': 'Y', 'GTT': 'V', 'GTC': 'V', 'GTA': 'V', 'GTG': 'V', 'TAG': '*', 'TGA': '*', 'TAA': '*'};
    this.iupacAmbiCodes = {'A': ['A'], 'C': ['C'], 'G': ['G'], 'T': ['T'], 'U': ['U'], 'M': ['A', 'C'], 'R': ['A', 'G'], 'W': ['A', 'T'], 'S': ['C', 'G'], 'Y': ['C', 'T'], 'K': ['G', 'T'], 'V': ['A', 'C', 'G'], 'H': ['A', 'C', 'T'], 'D': ['A', 'G', 'T'], 'B': ['C', 'G', 'T'], 'X': ['G', 'A', 'T', 'C'], 'N': ['G', 'A', 'T', 'C']};
}

TranslateDna.prototype = {
    constructor: TranslateDna,
    // Perform lookup to return an amino acid for a given nucleotide triplet.
    getAminoAcid: function (codon) {
        'use strict';
        return this.translateTable[codon];
    },
    // Longest and most complex method. Used to disambiguate mixed triplets.
    getCodonsFromAmbiguous: function (ambiguousCodon) {
        'use strict';
        var codons = [],
            first = this.iupacAmbiCodes[ambiguousCodon.charAt(0)],
            lenFirst = first.length,
            second = this.iupacAmbiCodes[ambiguousCodon.charAt(1)],
            lenSecond = second.length,
            third = this.iupacAmbiCodes[ambiguousCodon.charAt(2)],
            lenThird = third.length,
            nuc1,
            nuc2,
            nuc3,
            codon,
            i,
            j,
            k;

        for (i = 0; i < lenFirst; i += 1) {
            nuc1 = first[i];
            for (j = 0; j < lenSecond; j += 1) {
                nuc2 = second[j];
                for (k = 0; k < lenThird; k += 1) {
                    nuc3 = third[k];
                    codon = nuc1 + nuc2 + nuc3;
                    codons.push(codon);
                }
            }
        }
        return codons;
    },
    // Break the instance DNA sequence string into triplets (codons).  Assumes the DNA is in-frame.
    splitSequenceIntoTriplets: function () {
        'use strict';
        var i = 0,
            seqLen = this.dnaSeq.length,
            triplets = [],
            triplet;

        for (i = 0; i < seqLen; i += 3) {
            triplet = this.dnaSeq.slice(i, 3 + i);
            triplets.push(triplet);
        }
        this.triplets = triplets;
    },
    //Return an array of triplets, if the instance is set, return it, else generate the triplets array and return it.
    getTriplets: function () {
        'use strict';
        if (!this.triplets) {
            this.splitSequenceIntoTriplets();
        }
        return this.triplets;
    },
    getAminoAcids: function () {
        'use strict';
        var triplets = this.getTriplets(),
            tripletCount = triplets.length,
            aminoAcids = [],
            aminoAcid,
            mixedTriplets = [],
            mixedAminoAcids = [],
            i,
            j;

        for (i = 0; i < tripletCount; i += 1) {
            //Match only triplets composed of the four standard DNA nucleotide bases.
            if (triplets[i].match(/[ACGT]{3}/)) {
                aminoAcid = this.translateTable[triplets[i]];
                aminoAcids.push(aminoAcid);
            } else {
                //Allowable characters in input (four standard nucleotides A,C,G,T and all recognized IUPAC mixture codes).
                if (triplets[i].match(/[ACGTUMRWSYKVHDBXN]{3}/)) {
                    mixedTriplets = this.getCodonsFromAmbiguous(triplets[i]);
                    for (j = 0; j < mixedTriplets.length; j += 1) {
                        aminoAcid = this.translateTable[mixedTriplets[j]];
                        if (mixedAminoAcids.indexOf(aminoAcid) === -1) {
                            mixedAminoAcids.push(aminoAcid);
                        }
                    }
                    aminoAcids.push(mixedAminoAcids);
                } else {
                    aminoAcids.push(this.UNKNOWN);
                }
            }
        }
        return aminoAcids;
    }
};
