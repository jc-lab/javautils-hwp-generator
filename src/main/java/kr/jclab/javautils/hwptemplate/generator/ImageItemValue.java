package kr.jclab.javautils.hwptemplate.generator;

import lombok.AllArgsConstructor;
import lombok.Getter;

@AllArgsConstructor
public class ImageItemValue implements ItemValue {
    @Getter
    private final String ext;
    @Getter
    private final byte[] data;
}
