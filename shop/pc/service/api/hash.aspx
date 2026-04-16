<%@ Page Language="C#" ValidateRequest="false" Debug="false" %>
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.Text" %>
<%@ Import Namespace="System.Security.Cryptography" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Web" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Net" %>

<script runat="server">
/* 
 * Password Hashing With PBKDF2 (http://crackstation.net/hashing-security.htm).
 * Copyright (c) 2013, Taylor Hornby
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without 
 * modification, are permitted provided that the following conditions are met:
 *
 * 1. Redistributions of source code must retain the above copyright notice, 
 * this list of conditions and the following disclaimer.
 *
 * 2. Redistributions in binary form must reproduce the above copyright notice,
 * this list of conditions and the following disclaimer in the documentation 
 * and/or other materials provided with the distribution.
 *
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" 
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE 
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE 
 * ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE 
 * LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR 
 * CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF 
 * SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS 
 * INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN 
 * CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) 
 * ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE 
 * POSSIBILITY OF SUCH DAMAGE.
 */

    /// <summary>
    /// Salted password hashing with PBKDF2-SHA1.
    /// Author: havoc AT defuse.ca
    /// www: http://crackstation.net/hashing-security.htm
    /// Compatibility: .NET 3.0 and later.
    /// </summary>
    public class PasswordHash
    {
        // The following constants may be changed without breaking existing hashes.
        public const int SALT_BYTE_SIZE = 24;
        public const int HASH_BYTE_SIZE = 24;
        public const int PBKDF2_ITERATIONS = 20000;

        public const int ITERATION_INDEX = 0;
        public const int SALT_INDEX = 1;
        public const int PBKDF2_INDEX = 2;

	public const string PC_PREFIX="NSPC";
        /// <summary>
        /// Creates a salted PBKDF2 hash of the password.
        /// </summary>
        /// <param name="password">The password to hash.</param>
        /// <returns>The hash of the password.</returns>
        public static string CreateHash(string password)
        {
            // Generate a random salt
            RNGCryptoServiceProvider csprng = new RNGCryptoServiceProvider();
            byte[] salt = new byte[SALT_BYTE_SIZE];
            csprng.GetBytes(salt);

            // Hash the password and encode the parameters
            byte[] hash = PBKDF2(password, salt, PBKDF2_ITERATIONS, HASH_BYTE_SIZE);
            return PC_PREFIX + ":" +
                Convert.ToBase64String(salt) + ":" +
                Convert.ToBase64String(hash);
        }

        /// <summary>
        /// Validates a password given a hash of the correct one.
        /// </summary>
        /// <param name="password">The password to check.</param>
        /// <param name="correctHash">A hash of the correct password.</param>
        /// <returns>True if the password is correct. False otherwise.</returns>
        public static bool ValidatePassword(string password, string correctHash)
        {
            // Extract the parameters from the hash
            char[] delimiter = { ':' };
            string[] split = correctHash.Split(delimiter);
            int iterations = PBKDF2_ITERATIONS;
			string saveiter = split[0];
			if (saveiter!=PC_PREFIX) return(false);
            byte[] salt = Convert.FromBase64String(split[SALT_INDEX]);
            byte[] hash = Convert.FromBase64String(split[PBKDF2_INDEX]);

            byte[] testHash = PBKDF2(password, salt, iterations, hash.Length);
            return SlowEquals(hash, testHash);
        }

        /// <summary>
        /// Compares two byte arrays in length-constant time. This comparison
        /// method is used so that password hashes cannot be extracted from
        /// on-line systems using a timing attack and then attacked off-line.
        /// </summary>
        /// <param name="a">The first byte array.</param>
        /// <param name="b">The second byte array.</param>
        /// <returns>True if both byte arrays are equal. False otherwise.</returns>
        private static bool SlowEquals(byte[] a, byte[] b)
        {
            uint diff = (uint)a.Length ^ (uint)b.Length;
            for (int i = 0; i < a.Length && i < b.Length; i++)
                diff |= (uint)(a[i] ^ b[i]);
            return diff == 0;
        }

        /// <summary>
        /// Computes the PBKDF2-SHA1 hash of a password.
        /// </summary>
        /// <param name="password">The password to hash.</param>
        /// <param name="salt">The salt.</param>
        /// <param name="iterations">The PBKDF2 iteration count.</param>
        /// <param name="outputBytes">The length of the hash to generate, in bytes.</param>
        /// <returns>A hash of the password.</returns>
        private static byte[] PBKDF2(string password, byte[] salt, int iterations, int outputBytes)
        {
            Rfc2898DeriveBytes pbkdf2 = new Rfc2898DeriveBytes(password, salt);
            pbkdf2.IterationCount = iterations;
            return pbkdf2.GetBytes(outputBytes);
        }
    }
    public static bool IsLocalIpAddress(string host)
    {
        try
        { 
            IPAddress[] hostIPs = Dns.GetHostAddresses(host);  
            IPAddress[] localIPs = Dns.GetHostAddresses(Dns.GetHostName());
            foreach (IPAddress hostIP in hostIPs)
            {
                if (IPAddress.IsLoopback(hostIP)) return true;
                foreach (IPAddress localIP in localIPs)
                {
                    if (hostIP.Equals(localIP)) return true;
                }
            }
        }
        catch { }
        return false;
    }
</script>
<%


   // :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   // The code above is from: https://crackstation.net/hashing-security.htm#aspsourcecode
   // and is used according to the terms shown/included above
   // :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

   // :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   //
   // Code below this line is specific to ProductCart (by Netsource Commerce, Inc.)
   // Created for ProductCart and Netsource Commerce, Inc. by Michael Shaffer of
   // Amalgamated Switch & Signal Company, LLC
   // Rights to use, modify, publish this code in any way are granted to Netsource 
   // Commerce, Inc. by Michael Shaffer
   //
   // :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   //
   // THIS SOFTWARE IS PROVIDED BY MICHAEL SHAFFER (CONTRIBUTOR) "AS IS" 
   // AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE 
   // IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE 
   // ARE DISCLAIMED. IN NO EVENT SHALL THE CONTRIBUTOR OR HIS INTERESTS BE 
   // LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR 
   // CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF 
   // SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS 
   // INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN 
   // CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) 
   // ARISING IN ANY WAY OUT OF THE USE OF OR INABILITY TO USE THIS SOFTWARE, 
   // EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
   //
   // :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   //
   // NOTES:
   //
   //    The code checks to see that the request is HTTP POST, and
   //    that it originated from the local IP address. However, IP
   //    addresses can be spoofed. It is HIGHLY recommended that 
   //    the server have SSL enabled and that the requests made to
   //    this script use HTTPS. Possible enhancement: Pass a unique
   //    token with the HTTP POST (something that only this script 
   //    would know how to replicate) and validate it.
   //    
   //    There are only two actions (methods) to this script (passed
   //    in the parameter "AC"):
   //        G     = Generate a hash value for the supplied password ("reqPW")
   //        T     = Validate the supplied password ("reqPW") against the 
   //                supplied hash ("hash"). Returns either "True" or "False"
   //
   // :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::



   // Check to be sure that request is coming from local
   // If not, then show 'Err'. 
    if (!Request.IsLocal)
    {
        if (!IsLocalIpAddress(Request.ServerVariables["remote_addr"]))
        {
            // Response.Write("Err");
            // Response.End();
        }
    }


   // Get the request action
   string fn    = Request.Form["ac"];
   string reqPW = Request.Form["reqPW"];

   switch (fn)
   {
      // is the user requesting a password to be hashed?
      // if so, return the hash/salt string in this form:
      //   nnnnn:ssssss...:hhhhhh...
      // where nnnnn is the iteration count, ssssss is the
      // salt and hhhhhh is the hash
      case "G":
         if (reqPW.Length==0) {
            Response.Write("Err");
            Response.End();
         } else {      
            Response.Write(PasswordHash.CreateHash(reqPW));
            Response.End();
         }
         break;

      // is the user passing a password and hash to validate?
      // if so, use the passed 'nnnnn:ssssss...:hhhhhh...' string
      // and see if it's valid.
      case "T":   
         string hash = Request.Form["hash"];

         if (reqPW.Length==0 || hash.Length<5) {
            Response.Write("Err");
            Response.End();
         } else {      
            bool Passed = PasswordHash.ValidatePassword(reqPW, hash);
            Response.Write(Passed);
            Response.End();
         }
         break;

      default:
         Response.Write("Err");
         Response.End();
         break;
   }
%>