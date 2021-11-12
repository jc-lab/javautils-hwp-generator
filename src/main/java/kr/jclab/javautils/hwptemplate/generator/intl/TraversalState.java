package kr.jclab.javautils.hwptemplate.generator.intl;

import kr.dogfoot.hwplib.object.HWPFile;
import kr.jclab.javautils.hwptemplate.generator.GetItemValueHandler;
import kr.jclab.javautils.hwptemplate.generator.ItemValue;

import java.util.Deque;

@lombok.Builder
@lombok.Getter
public class TraversalState<ItemType> {
    private final boolean generateMode;

    private final HWPFile targetHwpFile;

    private final Deque<ItemType> items;
    private final GetItemValueHandler<ItemType> getItemValueHandler;

    private int templateCount = 0;
    private ItemType currentItem;

    public int getRemaining() {
        return this.items.size();
    }

    public boolean begin() {
        this.templateCount++;
        if (!this.isGenerateMode()) {
            return true;
        }
        if (this.items.isEmpty()) return false;
        this.currentItem = this.items.removeFirst();
        return true;
    }

    public boolean end() {
        if (!this.isGenerateMode()) {
            return true;
        }
        this.currentItem = null;
        return !this.items.isEmpty();
    }

    public ItemValue getItemValue(String key) {
        return this.getItemValueHandler.get(this.currentItem, key);
    }
}
