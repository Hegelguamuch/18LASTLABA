/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package proyecto1;
import java.util.Scanner;
/**
 *
 * @author estudiante
 */
public class Proyecto1 {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        int A;
        
        System.out.println("ingrese un numero");
        Scanner D1 = new Scanner(System.in);
        A = D1.nextInt();
        
        if (A%2==0) {
             System.out.println("es numero par");
        } else {
            System.out.println("No es numero par");
        }
        // TODO code application logic here
    }
    
}
