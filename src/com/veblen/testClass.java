package com.veblen;

/**
 * Created by dingyunxiang on 16/5/2.
 */
public class testClass {
    private int name;
    private int age;

    public int getName() {
        return name;
    }

    public void setName(int name) {
        this.name = name;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public testClass(int name, int age) {
        this.name = name;
        this.age = age;
    }

    public testClass() {
    }
}
