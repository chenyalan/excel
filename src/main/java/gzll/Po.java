package gzll;

import jdk.nashorn.internal.objects.annotations.Getter;
import jdk.nashorn.internal.objects.annotations.Setter;

/**
 * Created by 陈亚兰 on 2018/3/29.
 */
public class Po {
    private int start;
    private int end;
    public Po(){}
    public Po(int start,int end){
        this.start=start;
        this.end=end;
    }

    public int getStart() {
        return start;
    }

    public void setStart(int start) {
        this.start = start;
    }

    public int getEnd() {
        return end;
    }

    public void setEnd(int end) {
        this.end = end;
    }
}
