package pl.jsolve.templ4docx.util;

import java.util.List;

import pl.jsolve.sweetener.collection.Collections;

public class Key {

    private String key;
    private VariableType variableType;
    private List<Key> subKeys;

    public Key(String key, VariableType variableType) {
        this.key = key;
        this.variableType = variableType;
        this.subKeys = Collections.newArrayList();
    }

    public String getKey() {
        return key;
    }

    public VariableType getVariableType() {
        return variableType;
    }

    public void addSubKey(Key subkey) {
        this.subKeys.add(subkey);
    }

    public Key getFirstSubKey() {
        return subKeys.get(0);
    }

    public boolean containsSubKey() {
        return !subKeys.isEmpty();
    }

    public List<Key> getSubKeys() {
        return subKeys;
    }

    @Override
    public String toString() {
        return "Key [key=" + key + ", variableType=" + variableType + ", subKeys=" + subKeys + "]";
    }

}
