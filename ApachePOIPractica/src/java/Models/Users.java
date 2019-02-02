/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Models;

/**
 *
 * @author Danyer
 */
public class Users {
    public String password;
    public String email;
    public String userName;


public Users(String name, String email, String userName) {
    this.password = name;
    this.email = email;
    this.userName = userName;
}

        public String getPassword()
        {
                return password;
        }
        public void setPassword(String name)
        {
                this.password = name;
        }

        public String getEmail()
        {
                return email;
        }
        public void setEmail(String email)
        {
                this.email = email;
        }

        public String getUserName()
        {
                return userName;
        }
        public void setUserName(String userName)
        {
                this.userName = userName;
        }

}
