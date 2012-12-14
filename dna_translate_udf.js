/*global TranslateDna: false */
/*
 Follwing functions require the script file "TranslateDna.js".
 They can be called in Google spreadsheets as user-defined functions 
 Example Invocations as user-defined functions:
 =getTripletsFromDNA(A1, "|")
 =getAminoAcidsFromDNA(A1,"|")).

Code has been checked in JSLint
*/
// Concatenate the result of breaking up the input sequence into triplets with
// a user specified concatenation string and return the concatenated output.
function getTripletsFromDNA(dnaSeq, concatStr) {
    'use strict';
    var trans = new TranslateDna(dnaSeq);
    return trans.getTriplets().join(concatStr);
}
// Return the amino acid translation of the input sequence as a concatenated string.
// The array of amino acids can contain inner arrays when a given nucleotide triplet
// contains ambiguous characters such as "R".  These array elements are detected by
// the "Array.isArray()" check and are processed separately by being concatenated with an empty
// string and are converted to lower case.
function getAminoAcidsFromDNA(dnaSeq, concatStr) {
    'use strict';
    var trans = new TranslateDna(dnaSeq),
        aminoAcids = trans.getAminoAcids(),
        elementCount = aminoAcids.length,
        i,
        mixedAminoAcids,
        processedAminoAcids = [];
    for (i = 0; i < elementCount; i += 1) {
        if (Array.isArray(aminoAcids[i])) {
            mixedAminoAcids = aminoAcids[i].join('').toLowerCase();
            processedAminoAcids.push(mixedAminoAcids);
        } else {
            processedAminoAcids.push(aminoAcids[i]);
        }
    }
    return processedAminoAcids.join(concatStr);
}
