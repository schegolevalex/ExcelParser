package com.schegolevalex.excel.ExcelParser;

import lombok.*;
import lombok.experimental.FieldDefaults;

@Getter
@Setter
@FieldDefaults(level = AccessLevel.PRIVATE)
@AllArgsConstructor
@ToString
public class PipeTemp {
    Double outerDiameter;
    Double wallThickness;
    Double mass;
}
