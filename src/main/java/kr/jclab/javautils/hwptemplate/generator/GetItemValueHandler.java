package kr.jclab.javautils.hwptemplate.generator;

@FunctionalInterface
public interface GetItemValueHandler<ItemType> {
    ItemValue get(ItemType item, String key);
}
