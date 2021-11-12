package kr.jclab.javautils.hwptemplate.generator;

import lombok.AllArgsConstructor;
import lombok.Getter;

@AllArgsConstructor
public class TextItemValue implements ItemValue {
    @Getter
    private final String text;
}
