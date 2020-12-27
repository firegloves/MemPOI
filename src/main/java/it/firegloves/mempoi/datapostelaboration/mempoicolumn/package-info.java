/**
 * These classes supply a pipeline to apply any transformation on the generated export report.
 * If you don't know the pipeline pattern this could be a good time to discover it on the web.
 *
 * The concept is to have a pipeline of transformations that will be applied in sequence, as specified by the user when he
 * creates the MemPOI instance.
 *
 * Each pipeline step collects data during the report generation.
 * At the end  of the generation, each step applies its transformations to the workbook (excel report) resulting from the previous step.
 *
 * An example is the merged regions step. YOu can find a streamed API and a NON streamed API examples
 */
package it.firegloves.mempoi.datapostelaboration.mempoicolumn;
