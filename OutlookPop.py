import uuid
import requests
import ua_generator
import re
import traceback
import sys
import time  # Add time module for timing calculation

# Import logger
try:
    from main import Logger
    # Create OutlookPop's own logger instance
    logger = Logger("outlook_pop.log")
except ImportError:
    # If import fails, create a simple logger
    import time
    class SimpleLogger:
        def __init__(self):
            pass
        def _log(self, prefix, message):
            # Only log to file, no console output
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
            try:
                import os
                if not os.path.exists('logs'):
                    os.makedirs('logs')
                log_path = os.path.join('logs', 'outlook_pop.log')
                with open(log_path, 'a', encoding='utf-8') as f:
                    f.write(f"[{timestamp}] [{prefix}] {message}\n")
            except Exception:
                pass  # Silently ignore write errors
        def debug(self, message): self._log("DEBUG", message)
        def info(self, message): self._log("INFO", message)
        def success(self, message): self._log("SUCCESS", message)
        def warning(self, message): self._log("WARNING", message)
        def error(self, message): self._log("ERROR", message)
        def critical(self, message): self._log("CRITICAL", message)
    logger = SimpleLogger()

# Only output initialization info when running this file directly
if __name__ == "__main__":
    logger.info("=" * 50)
    logger.info("===== Outlook IMAP Service Activation Tool =====")
    logger.info("=" * 50)

class outlook:
    def __init__(self, ip):
        self.client = requests.session()
        # Initialize loginzt attribute to avoid exceptions
        self.loginzt = "unknown"
        
        # Configure proxy
        if ip and ip.strip():
            logger.info("[INIT] Using proxy: " + ip)
            # Format proxy address
            proxy_dict = {
                "http": f"http://{ip}",
                "https": f"http://{ip}"
            }
            self.client.proxies.update(proxy_dict)
            # No longer disable trust environment
            self.client.trust_env = True
        else:
            logger.info("[INIT] No proxy provided, using direct connection")
            self.client.proxies.clear()
            self.client.trust_env = False
        
        # Use fixed User-Agent, this is an older Chrome version that has proven higher success rate
        user_agent = "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.5481.154 Safari/537.36"
        self.headers = {
            "User-Agent": user_agent,
            "Accept-Language": "zh-CN,zh;q=0.9",
            "Accept-Encoding": "gzip, deflate, br, zstd",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
            "Upgrade-Insecure-Requests": "1",
            "Sec-Fetch-Site": "same-origin",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-User": "?1",
            "Sec-Fetch-Dest": "document"
        }
        
        logger.debug(f"[INIT] User-Agent type: {type(self.headers['User-Agent'])}")
        logger.debug(f"[INIT] User-Agent value: {self.headers['User-Agent']}")
    
    def login(self, username, password):
        try:
            logger.info(f"[LOGIN] Starting login for {username}")
            text = self.client.get(
                f"https://outlook.live.com/owa/?cobrandid={uuid.uuid4()}&nlp=1&deeplink=owa/",
                headers=self.headers
            ).text
            
            match = re.search('value="([^"]*)"', text)
            match2 = re.search("urlPost:'([^']*)'", text)
            
            if match and match2:
                PPFT = match.group(1)
                self.urlPost = match2.group(1)
                
                logger.debug(f"[LOGIN] Extracted PPFT: {PPFT[:20]}...")
                logger.debug(f"[LOGIN] Extracted urlPost: {self.urlPost}")
                
                data = {
                    "ps": 2,
                    "psRNGCDefaultType": "",
                    "psRNGCEntropy": "",
                    "psRNGCSLK": "",
                    "canary": "",
                    "ctx": "",
                    "hpgrequestid": "",
                    "PPFT": PPFT,
                    "PPSX": "Passp",
                    "NewUser": 1,
                    "FoundMSAs": "",
                    "fspost": 0,
                    "i21": 0,
                    "CookieDisclosure": 0,
                    "IsFidoSupported": 1,
                    "isSignupPost": 0,
                    "isRecoveryAttemptPost": 0,
                    "i13": 0,
                    "login": username,
                    "loginfmt": username,
                    "type": 11,
                    "LoginOptions": 3,
                    "lrt": "",
                    "lrtPartition": "",
                    "hisRegion": "",
                    "hisScaleUnit": "",
                    "passwd": password
                }
                
                logger.info(f"[LOGIN] Submitting login request...")
                response = self.client.post(
                    self.urlPost,
                    data=data,
                    headers=self.headers
                )
                
                logger.info(f"[LOGIN] Login response status code: {response.status_code}")
                
                text = response.text
                
                urlPost = re.search("urlPost:'([^']*)'", text)
                sFT = re.search("sFT:'([^']*)'", text)
                
                if urlPost and sFT:
                    self.loginzt = "owa"
                    self.urlPost2 = urlPost.group(1)
                    self.sFT = sFT.group(1)
                    logger.info(f"[LOGIN] Identified as OWA flow, obtained second stage URL and token")
                    return True
                
                action = re.search('action="([^"]*)"', text)
                
                if action:
                    self.action = action.group(1)
                    logger.info(f"[LOGIN] Identified as Action flow: {self.action}")
                    
                    values = re.findall('value="([^"]*)"', text)
                    if len(values) >= 3:
                        self.pprid = values[0]
                        self.ipt = values[1]
                        self.uaid = values[2]
                        logger.info(f"[LOGIN] Extracted pprid, ipt, uaid values")
                        
                        m = re.search(r'/([^/?]*)\?', self.action)
                        
                        if m:
                            self.loginzt = m.group(1)
                            logger.info(f"[LOGIN] Set login status to: {self.loginzt}")
                            return True
                    else:
                        logger.warning(f"[LOGIN] Not enough form values extracted, found {len(values)} values")
                        return False
                
                logger.warning(f"[LOGIN] Unrecognized login response, neither OWA token nor Action form found")
                return False
            
            logger.warning(f"[LOGIN] Initial page format unexpected, unable to extract PPFT or urlPost")
            self.loginzt = "error"
            return True
        
        except Exception as e:
            logger.error(f"[LOGIN] Login process error: {str(e)}")
            traceback.print_exc()
            return False
    
    def login2(self):
        try:
            logger.info("[LOGIN2] Starting second stage login...")
            
            data = {
                "LoginOptions": 3,
                "type": 28,
                "ctx": "",
                "hpgrequestid": "",
                "PPFT": self.sFT,
                "canary": ""
            }
            
            logger.info("[LOGIN2] Request headers and data:")
            logger.info(f"headers: {self.headers}")
            logger.info(f"data: {data}")
            logger.info(f"urlPost2: {self.urlPost2}")
            
            try:
                text = self.client.post(
                    self.urlPost2,
                    data=data,
                    headers=self.headers
                ).text
                
                fmHF_match = re.search('action="([^"]*)"', text)
                if not fmHF_match:
                    logger.warning("[LOGIN2] fmHF form action not found")
                    return False
                    
                fmHF = fmHF_match.group(1)
                logger.info(f"[LOGIN2] Found fmHF form action: {fmHF}")
                
                values = re.findall('value="([^"]*)"', text)
                logger.info(f"[LOGIN2] Found {len(values)} values")
                
                if len(values) < 6:
                    logger.warning(f"[LOGIN2] Not enough values extracted, need 6, only have {len(values)}")
                    return False
                
                wbids, pprid, wbid, NAP, ANON, t = values[:6]
                
                data = {
                    "wbids": wbids,
                    "pprid": pprid,
                    "wbid": wbid,
                    "NAP": NAP,
                    "ANON": ANON,
                    "t": t
                }
                
                logger.info(f"[LOGIN2] Submitting fmHF form data: {data}")
                
                ret = self.client.post(
                    fmHF,
                    data=data,
                    headers=self.headers
                )
                
                logger.info(f"[LOGIN2] Final URL: {ret.url}")
                
                canary_cookie = self.client.cookies.get("X-OWA-CANARY", path="/mail/0/")
                logger.info(f"[LOGIN2] X-OWA-CANARY Cookie: {canary_cookie}")
                
                if canary_cookie:
                    logger.success("[LOGIN2] Successfully obtained OWA canary cookie, login successful")
                    return True
                
                logger.warning("[LOGIN2] Failed to obtain OWA canary cookie, login failed")
                return False
            
            except Exception as e:
                logger.error(f"[LOGIN2] Form submission error: {str(e)}")
                traceback.print_exc()
                return False
                
        except Exception as e:
            logger.error(f"[LOGIN2] Second stage login error: {str(e)}")
            traceback.print_exc()
            return False
    
    def open_imap(self):
        try:
            logger.info("[IMAP] Starting IMAP service activation...")
            
            url = "https://outlook.live.com/owa/0/service.svc?action=SetConsumerMailbox&app=Mail&n=0"
            
            self.headers["action"] = "SetConsumerMailbox"
            canary = self.client.cookies.get("X-OWA-CANARY", path="/mail/0/")
            if not canary:
                logger.warning("[IMAP] Unable to get X-OWA-CANARY Cookie, IMAP activation failed")
                return False
                
            self.headers["x-owa-canary"] = canary
            self.headers["x-owa-urlpostdata"] = "%7B%22__type%22%3A%22SetConsumerMailboxRequest%3A%23Exchange%22%2C%22Header%22%3A%7B%22__type%22%3A%22JsonRequestHeaders%3A%23Exchange%22%2C%22RequestServerVersion%22%3A%22V2018_01_08%22%2C%22TimeZoneContext%22%3A%7B%22__type%22%3A%22TimeZoneContext%3A%23Exchange%22%2C%22TimeZoneDefinition%22%3A%7B%22__type%22%3A%22TimeZoneDefinitionType%3A%23Exchange%22%2C%22Id%22%3A%22Greenwich%20Standard%20Time%22%7D%7D%7D%2C%22Options%22%3A%7B%22PopEnabled%22%3Atrue%2C%22PopMessageDeleteEnabled%22%3Afalse%2C%22ImapEnabled%22%3Atrue%7D%7D"
            
            logger.info(f"[IMAP] Request headers: {self.headers}")
            
            try:
                ret = self.client.post(
                    url,
                    headers=self.headers
                )
                
                logger.info(f"[IMAP] IMAP activation response status code: {ret.status_code}")
                
                if "Header" in ret.text:
                    logger.success("[IMAP] Response contains Header, IMAP activation successful")
                    return True
                
                logger.warning(f"[IMAP] Response does not contain Header, IMAP activation failed")
                return False
            
            except Exception as e:
                logger.error(f"[IMAP] IMAP activation request error: {str(e)}")
                traceback.print_exc()
                return False
        
        except Exception as e:
            logger.error(f"[IMAP] IMAP activation process error: {str(e)}")
            traceback.print_exc()
            return False
    
    def Set_forwarding(self, email):
        url = "https://outlook.live.com/owa/0/service.svc?action=NewInboxRule&app=Mail&n=0"
        self.headers["action"] = "NewInboxRule"
        self.headers["x-owa-canary"] = self.client.cookies.get("X-OWA-CANARY", path="/mail/0/")
        self.headers["x-owa-urlpostdata"] = f'%7B%22__type%22%3A%22NewInboxRuleRequest%3A%23Exchange%22%2C%22Header%22%3A%7B%22__type%22%3A%22JsonRequestHeaders%3A%23Exchange%22%2C%22RequestServerVersion%22%3A%22V2018_01_08%22%2C%22TimeZoneContext%22%3A%7B%22__type%22%3A%22TimeZoneContext%3A%23Exchange%22%2C%22TimeZoneDefinition%22%3A%7B%22__type%22%3A%22TimeZoneDefinitionType%3A%23Exchange%22%2C%22Id%22%3A%22Greenwich%20Standard%20Time%22%7D%7D%7D%2C%22InboxRule%22%3A%7B%22Name%22%3A%22Forward%22%2C%22ForwardTo%22%3A%5B%7B%22__type%22%3A%22PeopleIdentity%3A%23Exchange%22%2C%22DisplayName%22%3A%22{email}%22%2C%22SmtpAddress%22%3A%22{email}%22%2C%22RoutingType%22%3A%22SMTP%22%7D%5D%2C%22StopProcessingRules%22%3Afalse%7D%7D'
        
        logger.info(self.headers)
        
        try:
            ret = self.client.post(
                url,
                headers=self.headers
            )
            
            logger.info(ret.text)
            
            if "Header" in ret.text:
                return True
            
            return False
        
        except:
            return False
    
    def login_zt(self):
        data = {
            "pprid": self.pprid,
            "ipt": self.ipt,
            "uaid": self.uaid
        }
        
        try:
            return self.client.post(
                self.action,
                data=data,
                headers=self.headers
            ).text
        
        except:
            return ""

def open_imap(uname, upwd, proxy, retries=10):
    logger.info(f"Starting IMAP service activation for account {uname}")
    
    if proxy and proxy.strip():
        logger.info(f"Using proxy: {proxy}")
    else:
        logger.info("No proxy provided, using direct connection")
    
    logger.info(f"Maximum attempts: {retries} times")
    
    start_time = time.time()  # Record start time
    
    for i in range(retries):
        try:
            logger.info(f"Attempt {i+1}/{retries} to activate IMAP")
            # Use provided proxy parameter
            oul = outlook(proxy)
            
            login_result = oul.login(uname, upwd)
            logger.info(f"Login result: {login_result}, Login status: {oul.loginzt}")
            
            if not login_result:
                logger.warning(f"Login failed, waiting 3 seconds before attempt {i+2}")
                time.sleep(3)  # Wait after failure before retry
                continue
                
            if oul.loginzt != "owa":
                logger.warning(f"Login status is not owa, but: {oul.loginzt}, cannot continue IMAP activation")
                if hasattr(oul, 'action'):
                    logger.info(f"May need other processing: action={oul.action}")
                logger.warning(f"Waiting 3 seconds before attempt {i+2}")
                time.sleep(3)  # Wait after failure before retry
                continue
                
            logger.info("Starting second stage login")
            login2_result = oul.login2()
            logger.info(f"Second stage login result: {login2_result}")
            
            if not login2_result:
                logger.warning(f"Second stage login failed, waiting 3 seconds before attempt {i+2}")
                time.sleep(3)  # Wait after failure before retry
                continue
                
            logger.info("Requesting IMAP service activation")
            open_imap_result = oul.open_imap()
            logger.info(f"IMAP service activation result: {open_imap_result}")
            
            if open_imap_result:
                end_time = time.time()  # Record end time
                elapsed_time = end_time - start_time  # Calculate time taken
                logger.success(f"Successfully activated IMAP service for account {uname}")
                logger.success(f"Success on attempt {i+1}, total time: {elapsed_time:.2f} seconds")
                return True
            else:
                logger.warning(f"IMAP service activation request failed, waiting 3 seconds before attempt {i+2}")
                time.sleep(3)  # Wait after failure before retry
                
        except Exception as e:
            logger.error(f"IMAP activation process error: {str(e)}")
            traceback.print_exc()
            
            # If not the last attempt, wait and continue
            if i < retries - 1:
                logger.warning(f"Waiting 3 seconds before attempt {i+2}")
                time.sleep(3)  # Wait after error before retry
    
    end_time = time.time()  # Record end time (failure case)
    elapsed_time = end_time - start_time  # Calculate time taken
    logger.error(f"Failed to activate IMAP service for account {uname}")
    logger.error(f"Tried {retries} times, total time: {elapsed_time:.2f} seconds")
    return False

def main():
    """Main function - silent mode"""
    # Process command line arguments
    if len(sys.argv) > 2:
        # Get account and password from command line arguments
        email = sys.argv[1]
        password = sys.argv[2]
        proxy = sys.argv[3] if len(sys.argv) > 3 else ""
        
        logger.info(f"Using command line arguments - Account: {email}")
    else:
        logger.error("Error: Insufficient parameters provided, please provide account and password")
        logger.info("Usage: python OutlookPop.py <email> <password> [proxy]")
        return 1
    
    # Start IMAP activation
    logger.info("Starting processing...")
    start_time = time.time()  # Record total processing start time
    result = open_imap(email, password, proxy)
    end_time = time.time()  # Record total processing end time
    total_time = end_time - start_time  # Calculate total time taken
    
    # Record result
    if result:
        logger.success(f"Account {email} IMAP service activated")
    else:
        logger.error(f"Failed to activate IMAP service for account {email}")
    logger.info(f"Total time: {total_time:.2f} seconds")
    
    return 0 if result else 1

if __name__ == "__main__":
    try:
        exit_code = main()
        sys.exit(exit_code)
    except KeyboardInterrupt:
        logger.info("Operation cancelled")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Program error: {str(e)}")
        traceback.print_exc()
        sys.exit(1) 