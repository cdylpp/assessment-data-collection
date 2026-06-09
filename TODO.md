# TODO

- [ ] Focus our time on generating the final master excel file. The program should compile one or more node templates, relate the scores back to each participant and aggregate scores based on aggregate metrics. The way the program will be used is to batch process a set of incomplete node templates. For example, after scoring the templates will be placed inside a drop box

```text
dropbox/
    soac_block1_recorder1.xlsx
    soac_block1_recorder2.xlsx
    soac_block1_recorder3.xlsx
    ...
```

The program will read the Excel files in the dropbox, and process each of the participants scores. We can assume that each file creates a partition on the set of candidates so that no two files have scores on the same candidates. We can then reconstruct / compile the scores for all participants and output the master roster than shows all evolutions in a single file. 

Once the files have been processed they should be moved to a `processed/` folder. We should build the program so that scores are compiled cummulatively. For example, suppose the duration of the assessment is a week long. Each night the graders drop their templates into the drop box (partially complete since some scores are for later in the week). The program will compile all scores that have been captured up to that evening. The next day the same will happen. By the end of the week all scores should be compiled for the master file.

## Backlog

- [ ] Define formal config schemas for config, metrics, evolutions, events, and master mappings. Blocked on schema/modeling dependencies.
- [ ] Introduce explicit strict and permissive generation modes. Blocked on schema/modeling dependencies.
