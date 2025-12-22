import java.util.*;
class Solution {
    public static int myAtoi(String s) {
        StringBuilder sb=new StringBuilder();
        for(int i=0;i<s.length();i++){
            if(s.charAt(i)!=' '){
                sb.append(s.charAt(i));
            }else if(sb.length()>0){
                sb.append(s.charAt(i));
            }
        }
        System.out.println(sb);
      boolean neg=false;
        for(int i=0;i<sb.length();i++){
            if(sb.charAt(i)=='-'){
                neg=true;
                sb.deleteCharAt(i);
                break;
            }else if(sb.charAt(i)=='+'){
                sb.deleteCharAt(i);
                break;
            }else{
              break;
            }
        }
        System.out.println(sb);
        for(int i=0;i<sb.length();i++){
            if(sb.charAt(i)=='0'){
                sb.deleteCharAt(i);

            }

        }
        System.out.println(sb);
         
        StringBuilder sb2=new StringBuilder();
        for(int i=0;i<sb.length();i++){
            if(sb.charAt(i)>='0'&&sb.charAt(i)<='9'){
                sb2.append(sb.charAt(i));
            }else{
                break;
            }

        }
        System.out.println(sb2);
        if(sb2.length()==0){
            return 0;
        }
        long v=Long.valueOf(sb2.toString());
        if(neg){
            v=v*-1;
        }
        if(v>Integer.MAX_VALUE){
            return Integer.MAX_VALUE;
        }else if(v<Integer.MIN_VALUE){
            return Integer.MIN_VALUE;
        }
        return (int)v;
    }
    public static void main(String[] args) {
      String s="0-1";
        System.out.println(myAtoi(s));
    }
}