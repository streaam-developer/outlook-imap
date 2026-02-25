import random
import threading
import time
from typing import Dict, List, Tuple
import uuid
import re
from datetime import datetime, timezone
from functools import cached_property
from urllib.parse import quote
import json
import os
import traceback
import string  # Add string module import
import sys
import urllib.parse
import urllib3

# Import OutlookPop module functionality
from OutlookPop import open_imap as enable_imap

# Check if OauthBlueMail module is available
try:
    # Use importlib to dynamically check module availability instead of direct import
    import importlib.util
    spec = importlib.util.find_spec("OauthBlueMail")
    BLUEMAIL_AVAILABLE = spec is not None
    
    if BLUEMAIL_AVAILABLE:
        print("[bold green]OauthBlueMail module detected, OAuth token functionality available[/bold green]")
    else:
        print("[bold yellow]Warning: OauthBlueMail module not found, OAuth token will not be obtained[/bold yellow]")
except Exception as e:
    print(f"[bold yellow]Error checking OauthBlueMail module: {str(e)}[/bold yellow]")
    BLUEMAIL_AVAILABLE = False

# Function merged from random_strings.py
def random_string(length=10, upper=True, lower=True, digit=True, special=False):
    chars = ""
    if upper:
        chars += string.ascii_uppercase
    if lower:
        chars += string.ascii_lowercase
    if digit:
        chars += string.digits
    if special:
        chars += string.punctuation
        
    if not chars:
        chars = string.ascii_letters + string.digits
        
    return ''.join(random.choice(chars) for _ in range(length))

import faker
from curl_cffi import requests
from rich import print

# Remove import line for random_strings
# from random_strings import random_string

# Add captcha API key configuration
# Track used API keys

# Get new KoCaptcha API key from file
# Proxy API configuration

# Enhanced logger class - unified format
class Logger:
    # Log level definitions
    LEVELS = {
        "DEBUG": "[blue]",
        "INFO": "[cyan]",
        "SUCCESS": "[green]",
        "WARNING": "[yellow]",
        "ERROR": "[red]",
        "CRITICAL": "[bold red]"
    }
    
    def __init__(self, log_file="outlook_tool.log"):
        self.log_file = log_file
        # Ensure log directory exists
        if not os.path.exists('logs'):
            os.makedirs('logs')
        self.log_path = os.path.join('logs', log_file)
        
        # Remove automatic startup log recording
        # self._log("INFO", "===== Microsoft Outlook Automatic Account Creation Tool =====")
    
    def _log(self, level, message):
        """Internal logging method, logs to both file and console"""
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        short_timestamp = time.strftime("%H:%M:%S")
        
        # Log to file
        try:
            with open(self.log_path, 'a', encoding='utf-8') as f:
                f.write(f"[{timestamp}] [{level}] {message}\n")
        except Exception as e:
            print(f"[bold red]Failed to write to log: {str(e)}[/bold red]")
        
        # Output to console (using rich format)
        color_code = self.LEVELS.get(level, "[white]")
        formatted_message = f"{color_code}[{short_timestamp}] {message}[/{color_code.strip('[]')}]"
        print(formatted_message)
        
    def debug(self, message):
        """Debug level log"""
        self._log("DEBUG", message)
    
    def info(self, message):
        """Info level log"""
        self._log("INFO", message)
    
    def success(self, message):
        """Success level log"""
        self._log("SUCCESS", message)
    
    def warning(self, message):
        """Warning level log"""
        self._log("WARNING", message)
    
    def error(self, message):
        """Error level log"""
        self._log("ERROR", message)
    
    def critical(self, message):
        """Critical error level log"""
        self._log("CRITICAL", message)
    
    # For compatibility with old code, keep original methods but redirect to new ones
    def file(self, message):
        """Log message to file (compatible with old method)"""
        self._log("INFO", message)
    
    def console(self, message):
        """Output message to console (compatible with old method)"""
        self._log("INFO", message)

# Create global logger
logger = Logger()

# Define supported Outlook suffixes and corresponding region parameters
OUTLOOK_DOMAINS = [
    # Format: (domain name, domain suffix, MKT parameter)
    ("outlook", "jp", "JA-JP"),        # Japan
    ("outlook", "kr", "KO-KR"),        # Korea
    ("outlook", "sa", "AR-AR"),        # Saudi Arabia
    ("outlook.com", "ar", "AU-AR"),    # Argentina
    ("outlook.com", "au", "AU-AU"),    # Australia
    ("outlook", "at", "AT-AT"),        # Austria
    ("outlook", "be", "BA-BE"),        # Belgium
    ("outlook.com", "br", "Br-Br"),    # Brazil
    ("outlook", "cl", "cl-cl"),        # Chile
    ("outlook", "cz", "cz-cz"),        # Czech Republic
    ("outlook", "fr", "fr-fr"),        # France
    ("outlook", "de", "de-de"),        # Germany
    ("outlook.com", "gr", "gr-gr"),    # Greece
    ("outlook.co", "il", "il-il"),     # Israel
    ("outlook", "in", "in-in"),        # India
    ("outlook.co", "id", "iD-iD"),     # Indonesia
    ("outlook", "ie", "ie-ie"),        # Ireland
    ("outlook", "it", "it-it"),        # Italy
    ("outlook", "hu", "hu-hu"),        # Hungary
    ("outlook", "lv", "lv-lv"),        # Latvia
    ("outlook", "my", "my-my"),        # Malaysia
    ("outlook.co", "nz", "nz-nz"),     # New Zealand
    ("outlook", "ph", "ph-ph"),        # Philippines
    ("outlook", "pt", "pt-pt"),        # Portugal
    ("outlook", "sg", "sg-sg"),        # Singapore
    ("outlook", "sk", "sk-sk"),        # Slovakia
    ("outlook", "es", "es-es"),        # Spain
    ("outlook.co", "th", "th-th"),     # Thailand
    ("outlook.com", "tr", "tr-tr"),    # Turkey
    ("outlook.com", "vn", "vn-vn"),    # Vietnam
    ("outlook", "dk", "da-dk"),        # Denmark
    # Add traditional domains
    ("outlook", "com", "en-US"),       # USA/International
    ("hotmail", "com", "en-US"),       # USA/International
]

def get_proxy_from_file() -> str:
    """
    Get proxy from local file
    
    Returns:
        Proxy in format ip:port:username:password, or None if failed
    """
    try:
        logger.info("Getting proxy from file...")
        with open("data/proxies.txt", "r") as f:
            proxies = f.read().splitlines()
            
        if not proxies:
            logger.error("No proxies found in file")
            return None
            
        # Get random proxy from list
        return random.choice(proxies)
    except Exception as e:
        logger.error(f"Error getting proxy from file: {str(e)}")
        return None

def check_proxy_available(proxy: str, timeout: int = 5) -> bool:
    """
    Check if proxy IP is available
    
    Args:
        proxy: Proxy IP in format ip:port
        timeout: Timeout in seconds, default 5
        
    Returns:
        bool: True if proxy is available, False otherwise
    """
    if not proxy or not proxy.strip():
        logger.warning("No proxy provided, cannot check availability")
        return False
        
    try:
        logger.info(f"Checking proxy availability: {proxy}")
        # Set proxy format
        proxies = {
            "http": f"http://{proxy}",
            "https": f"http://{proxy}"
        }
        
        # Use curl_cffi.requests library for requests, consistent with main program
        # Test connection to Microsoft login page as this is the target site for subsequent operations
        test_url = "https://login.live.com/login.srf"
        response = requests.get(
            test_url,
            proxies=proxies,
            timeout=timeout
        )
        
        # Check response status code
        if response.status_code == 200:
            logger.success(f"Proxy available: {proxy}")
            return True
        else:
            logger.warning(f"Proxy unavailable, status code: {response.status_code}")
            return False
            
    except Exception as e:
        logger.error(f"Error checking proxy availability: {str(e)}")
        return False

def solve_captcha_kocaptcha(arkose_blob, proxy: str, caller=None, retry_count=0):
    """
    Solve captcha using local service
    
    Args:
        arkose_blob: Captcha blob data
        proxy: Proxy IP in format ip:port
        caller: Caller object for setting status flags
        retry_count: Retry count for API key change, prevent infinite loop
        
    Returns:
        Captcha solution token, or None if failed
    """
    global logger
    
    # Prevent infinite retries
    if retry_count >= 5:
        logger.error(f"❌ Captcha solving failed: Reached maximum API key retry count")
        if caller and hasattr(caller, 'captcha_failed'):
            caller.captcha_failed = True
        return None
    
    # Save original log function for later restoration
    original_log_function = logger._log
    
    # Define temporary log function that only writes to file
    def log_to_file_only(level, message):
        """Only write to log file, no console output"""
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        try:
            with open(logger.log_path, 'a', encoding='utf-8') as f:
                f.write(f"[{timestamp}] [{level}] {message}\n")
        except Exception:
            pass  # Ignore file write errors
    
    try:
        # Replace log function to only write to file
        logger._log = log_to_file_only
        
        # Log to file (not shown in console)
        log_to_file_only("INFO", f"🔍 Using local service to solve captcha")
        
        # Create task
        try:
            response = requests.post(
                "http://127.0.0.1:5000/solve",
                json={
                    "blob": arkose_blob,
                    "private_key": "B7D8911C-5CC8-A9A3-35B0-554ACEE604DA",
                    "og_proxy": f"http://{proxy}"
                },
                timeout=120
            )
            
            log_to_file_only("DEBUG", f"Captcha task submission response: {response.text}")
            
            if not response.text.strip():
                log_to_file_only("ERROR", f"❌ Captcha service returned empty response")
                if caller and hasattr(caller, 'captcha_failed'):
                    caller.captcha_failed = True
                return None
                
            result = response.json()
            
            # Check for errors
            if result.get("error", False):
                error_code = result.get("error_code", "unknown")
                error_desc = result.get("error_description", "Unknown error")
                
                log_to_file_only("ERROR", f"❌ Captcha task creation failed: {error_code} - {error_desc}")
                if caller and hasattr(caller, 'captcha_failed'):
                    caller.captcha_failed = True
                return None
            
            token = result.get("token")
            
            if token:
                # Success status
                log_to_file_only("SUCCESS", f"✅ Captcha solved (token: {token[:15]}...{token[-15:] if len(token) > 30 else ''})")
                log_to_file_only("DEBUG", f"Captcha full token: {token}")
                # Restore original log function
                logger._log = original_log_function
                return token
            else:
                log_to_file_only("ERROR", f"❌ Captcha result format abnormal")
                log_to_file_only("DEBUG", f"Captcha result format abnormal: {json.dumps(result)}")
                # Restore original log function
                logger._log = original_log_function
                if caller and hasattr(caller, 'captcha_failed'):
                    caller.captcha_failed = True
                return None
                
        except requests.RequestException as e:
            log_to_file_only("ERROR", f"❌ Failed to create captcha task: {str(e)}")
            log_to_file_only("DEBUG", f"Detailed error creating captcha task: {traceback.format_exc()}")
            if caller and hasattr(caller, 'captcha_failed'):
                caller.captcha_failed = True
            return None
        except json.JSONDecodeError as e:
            log_to_file_only("ERROR", f"❌ Failed to parse captcha task response: {str(e)}")
            log_to_file_only("DEBUG", f"Detailed error parsing captcha task: {traceback.format_exc()}")
            if caller and hasattr(caller, 'captcha_failed'):
                caller.captcha_failed = True
            return None
            
    except Exception as e:
        # Log error but don't output to console
        log_to_file_only("ERROR", f"❌ Local service processing failed: {str(e)}")
        log_to_file_only("DEBUG", f"Detailed error in local service processing: {traceback.format_exc()}")
        
        # Restore original log function
        logger._log = original_log_function
        
        # Output brief error message and newline
        print(f"🔄 Waiting for captcha result... - Processing failed", end="\r", flush=True)
        print()
        
        if caller and hasattr(caller, 'captcha_failed'):
            caller.captcha_failed = True
        return None
        
    finally:
        # Ensure original log function is restored in all cases
        logger._log = original_log_function

class Outlook:
    def __init__(self) -> None:
        logger.info("[bold blue]Initializing Outlook class...[/bold blue]")
        self.faker = faker.Faker()
        self.should_gen = True
        self.captcha_failed = False  # Add captcha failure flag
        if not os.path.exists('out'):
            os.makedirs('out')
            
        # Print supported domains
        self._print_supported_domains()
        
    def _print_supported_domains(self):
        """Print all supported domains"""
        logger.info("[bold green]Supported Outlook domains:[/bold green]")
        for i, (name, tld, mkt) in enumerate(OUTLOOK_DOMAINS, 1):
            domain = f"{name}.{tld}"
            logger.info(f"[cyan]{i:2d}. {domain:<20} (Region: {mkt})[/cyan]")

    @cached_property
    def canary_re(self) -> re.Pattern:
        return re.compile(r'"apiCanary":"(.*)","iUiFlavor')

    @cached_property
    def hpgid_re(self) -> re.Pattern:
        return re.compile(r'"hpgid":(\d*)')

    @cached_property
    def scenario_id_re(self) -> re.Pattern:
        return re.compile(r'"iScenarioId":(\d*)')

    @cached_property
    def final_url_re(self) -> re.Pattern:
        return re.compile(r'"urlFinalBack":"(.*)&res=Cancel')
    
    @cached_property
    def urldfp_re(self) -> re.Pattern:
        return re.compile(r'"urlDfp":"(.*)","urlHipChallenge')

    @cached_property
    def hfid_re(self) -> re.Pattern:
        return re.compile(r'"sHipFid":"(\w*)')

    def get_base_params(self, mkt) -> Dict[str, str]:
        """Get base request parameters based on region parameter"""
        return {
            'haschrome': '1',
            'mkt': mkt,  # Use specific region parameter
            'client_info': '1',
            'scope': 'profile offline_access openid service::outlook.office.com::MBI_SSL',
            'signup': '1',
            'response_type': 'code',
            'fl': 'wld',
            'display': 'touch',
            'client_id': '00000000-0000-0000-0000-000048170EF2',
            'noauthcancel': '1',
            'claims': '{"compact":{"name":{"essential":true}}}',
            'redirect_uri': 'msauth://com.microsoft.office.outlook/fcg80qvoM1YMKJZibjBwQcDfOno%3D',
            'x-ms-sso-ignore-sso': '1'
        }

    @property
    def base_headers(self) -> Dict[str, str]:
        _uuid = str(uuid.uuid4())
        return {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
            'Accept-Encoding': 'gzip, deflate, br, zstd',
            'Accept-Language': 'en-US,en;q=0.9',
            'client-request-id': _uuid,
            'Connection': 'keep-alive',
            'correlation-id': _uuid,
            'Host': 'login.microsoftonline.com',
            'return-client-request-id': 'false',
            'sec-ch-ua': '"Android WebView";v="135", "Not-A.Brand";v="8", "Chromium";v="135"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Android"',
            'Sec-Fetch-Dest': 'document',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-Site': 'none',
            'Sec-Fetch-User': '?1',
            'Upgrade-Insecure-Requests': '1',
            'User-Agent': 'Mozilla/5.0 (Linux; Android 9; SM-S9210 Build/PQ3A.190605.07291528; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/135.0.7049.100 Safari/537.36 PKeyAuth/1.0',
            'x-client-os': '28',
            'x-client-sku': 'MSAL.xplat.android',
            'x-client-src-sku': 'MSAL.xplat.android',
            'x-client-ver': '1.1.0+0f7f6135',
            'x-ms-passkeyauth': '1.0/passkey',
            'x-ms-sso-ignore-sso': '1',
            'X-Requested-With': 'com.microsoft.office.outlook'
        }

    def init_session(self, domain_index: int = None) -> None:
        """
        Initialize session and create account
        
        Args:
            domain_index: Domain index to use (1-based), if None then random selection
        """
        while self.should_gen:
            try:
                print("[bold cyan]Starting to create new email account...[/bold cyan]")
                email = random_string(upper=False, length=8, digit=False)
                print(f"[cyan]Generated email prefix: {email}[/cyan]")
                
                # Select domain
                if domain_index is not None and 1 <= domain_index <= len(OUTLOOK_DOMAINS):
                    domain_choice = OUTLOOK_DOMAINS[domain_index - 1]
                else:
                    domain_choice = random.choice(OUTLOOK_DOMAINS)
                
                name, tld, mkt = domain_choice
                domain_full = f"{name}.{tld}"
                print(f"[cyan]Selected domain: {domain_full} (Region: {mkt})[/cyan]")
                
                session = requests.Session(impersonate="chrome110")
                session.headers = self.base_headers
                
                # Get proxy from API
                proxy = get_proxy_from_file()
                if proxy:
                    print(f"[cyan]Using API proxy: {proxy}[/cyan]")
                    proxies = {'http': f'http://{proxy}', 'https': f'http://{proxy}'}
                    session.proxies = proxies
                else:
                    print("[bold yellow]⚠️ Unable to get proxy, will use direct connection[/bold yellow]")

                print("[cyan]Accessing Microsoft login page...[/cyan]")
                try:
                    # Use region-specific parameters
                    r = session.get(
                        'https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize', 
                        params=self.get_base_params(mkt)
                    )
                    r.raise_for_status()
                except Exception as e:
                    print(f"[bold red]Error requesting Microsoft login page: {str(e)}[/bold red]")
                    print("[bold yellow]Restarting entire creation process...[/bold yellow]")
                    continue  # Restart entire process

                results: Dict[str, str] = {"canary_token": None, "hpgid": None, "scenario_id": None, "final_url": None, 'urldfp': None}

                text_buffer = r.text

                for param_name, regex in [
                    ("canary_token", self.canary_re),
                    ("hpgid", self.hpgid_re),
                    ("scenario_id", self.scenario_id_re),
                    ("final_url", self.final_url_re),
                    ("urldfp", self.urldfp_re),
                ]:
                    if results[param_name] is None:
                        match = regex.search(text_buffer)
                        if match:
                            if param_name == "canary_token":
                                results[param_name] = match.group(1).encode('utf-8').decode('unicode_escape')
                            else:
                                results[param_name] = match.group(1)
                            print(f"[green]Successfully extracted {param_name}[/green]")

                if all(results.values()):
                    print("[green]Successfully extracted all required parameters[/green]")
                else:
                    missing = [k for k, v in results.items() if v is None]
                    print(f"[bold red]Unable to extract required parameters: {', '.join(missing)}[/bold red]")
                    print("[bold yellow]Restarting entire creation process...[/bold yellow]")
                    continue  # Restart entire process

                
                session.headers.update(
                    {
                        'Accept': 'application/json',
                        'canary': results['canary_token'],
                        'client-request-id': session.headers['client-request-id'].replace('-', ''),
                        'Content-type': 'application/json; charset=utf-8',
                        'correlationId': session.headers['client-request-id'].replace('-', ''),
                        'Host': 'signup.live.com',
                        'hpgact': '0',
                        'hpgid': results['hpgid'],
                        'Origin': 'https://signup.live.com',
                        'Referer': r.url,
                        'Sec-Fetch-Dest': 'empty',
                        'Sec-Fetch-Mode': 'cors',
                        'Sec-Fetch-Site': 'same-origin'
                    }
                )
                print("[cyan]Updating session headers...[/cyan]")
                
                payload = {
                    "clientExperiments": [
                        {
                            "parallax": "allowflowidnavigation",
                            "control": "allowflowidnavigation_control",
                            "treatments": [
                                "allowflowidnavigation_treatment"
                            ]
                        },
                        {
                            "parallax": "addprivatebrowsingtexttofabricfooter",
                            "control": "addprivatebrowsingtexttofabricfooter_control",
                            "treatments": [
                                "addprivatebrowsingtexttofabricfooter_treatment"
                            ]
                        }
                    ]
                }
                print("[cyan]Sending experiment evaluation request...[/cyan]")
                try:
                    r2 = session.post('https://signup.live.com/API/EvaluateExperimentAssignments', json=payload)
                    r2.raise_for_status()
                except Exception as e:
                    print(f"[bold red]Error in experiment evaluation request: {str(e)}[/bold red]")
                    print("[bold yellow]Restarting entire creation process...[/bold yellow]")
                    continue  # Restart entire process
                
                api_canary = r2.json().get('apiCanary')
                
                if not api_canary:
                    print("[bold red]Unable to extract api canary[/bold red]")
                    print("[bold yellow]Restarting entire creation process...[/bold yellow]")
                    continue  # Restart entire process
                print("[green]Successfully extracted api canary[/green]")
                
                session.headers['canary'] = api_canary
                
                print("[cyan]Accessing DFP URL...[/cyan]")
                try:
                    session.get(results['urldfp'])
                except Exception as e:
                    print(f"[bold red]Error accessing DFP URL: {str(e)}[/bold red]")
                    print("[bold yellow]Restarting entire creation process...[/bold yellow]")
                    continue  # Restart entire process
                
                params = {
                    'sru': results['final_url'],
                    'mkt': mkt,  # Use region-specific parameter
                    'uiflavor': 'host',
                    'fl': 'wld',
                    'client_id': '00000000-0000-0000-0000-000048170EF2',
                    'noauthcancel': '1',
                    'uaid': session.headers['correlationId'],
                    'suc': '00000000-0000-0000-0000-000048170EF2'
                }
                
                password = random_string(16) + "?@!_A"
                print(f"[cyan]Creating account with domain: {domain_full} and random password[/cyan]")
                
                # Build availability check mapping list - only check the specific domain
                avail_check_list = [
                    f"{email}@{domain_full}:false"
                ]
                
                if tld == 'com':
                    country = 'US'
                else:
                    country = tld.upper()

                payload = {
                    "BirthDate": f"{random.randint(1,28)}:{random.randint(1,12)}:{random.randint(1970, 2005)}",
                    "CheckAvailStateMap": avail_check_list,
                    "Country": country,
                    "EvictionWarningShown": [],
                    "FirstName": self.faker.first_name(),
                    "IsRDM": False,
                    "IsOptOutEmailDefault": False,
                    "IsOptOutEmailShown": 1,
                    "IsOptOutEmail": False,
                    "IsUserConsentedToChinaPIPL": False,
                    "LastName": self.faker.last_name(),
                    "LW": 1,
                    "MemberName": f"{email}@{domain_full}",
                    "RequestTimeStamp": datetime.now(timezone.utc).isoformat(timespec='milliseconds') + "Z",
                    "ReturnUrl": "",
                    "SignupReturnUrl": quote(results['final_url']).replace('%3D', '%3d',).replace('%3A', ':'),
                    "SuggestedAccountType": "OUTLOOK",
                    "SiteId": "68692",
                    "VerificationCodeSlt": "",
                    "WReply": "",
                    "MemberNameChangeCount": 1,
                    "MemberNameAvailableCount": 1,
                    "MemberNameUnavailableCount": 0,
                    "Password": password,
                    "uiflvr": 1,
                    "scid": results['scenario_id'],
                    "uaid": session.headers['correlationId'],
                    "hpgid": results['hpgid']
                }
                
                print("[cyan]Submitting account creation request...[/cyan]")
                try:
                    r = session.post('https://signup.live.com/API/CreateAccount', params=params, json=payload)
                    print(f"[yellow]Account creation response: {r.status_code}[/yellow]")
                    
                    # Process API response
                    if r.status_code == 200:
                        response_json = r.json()
                        if 'error' in response_json:
                            error_code = response_json['error'].get('code')
                            error_field = response_json['error'].get('field')
                            
                            print(f"[yellow]Encountered error: Code {error_code}, Field {error_field}[/yellow]")
                            
                            # Special handling for 1041 error (requires captcha)
                            if error_code == "1041":
                                # Captcha handling logic remains unchanged
                                print("[yellow]Captcha solving required...[/yellow]")
                                account_creation_jsn = json.loads(r.json()['error']['data'])
                                arkose_blob = account_creation_jsn['arkoseBlob']
                                print(arkose_blob)
                                
                                # Use KoCaptcha service to solve captcha
                                self.captcha_failed = False  # Reset failure flag
                                captcha_token = solve_captcha_kocaptcha(arkose_blob, proxy, self)
                                
                                if not captcha_token or self.captcha_failed:
                                    print("[bold red]Captcha solve failed, cannot continue[/bold red]")
                                    print("[bold yellow]Restarting entire creation process...[/bold yellow]")
                                    continue  # Restart entire process
                                
                                print(f"[green][+] Successfully obtained captcha solution: {captcha_token[:16]}...[/green]")
                                
                                # Update payload to include captcha token
                                payload.update(
                                    {
                                        'RiskAssessmentDetails': account_creation_jsn['riskAssessmentDetails'],
                                        'RepMapRequestIdentifierDetails': account_creation_jsn['repMapRequestIdentifierDetails'],
                                        'HFId': None,
                                        'HPId': 'B7D8911C-5CC8-A9A3-35B0-554ACEE604DA',
                                        'HSol': captcha_token,
                                        'HType': 'enforcement',
                                        'HId': captcha_token,
                                        "uiflvr": 1,
                                        "scid": results['scenario_id'],
                                        "uaid": session.headers['correlationId'],
                                        "hpgid": results['hpgid']
                                    }
                                )
                                
                                print("[cyan]Submitting account creation request with captcha solution...[/cyan]")
                                try:
                                    r = session.post('https://signup.live.com/API/CreateAccount', params=params, json=payload)
                                    r.raise_for_status()
                                except Exception as e:
                                    print(f"[bold red]Error in account creation request with captcha: {str(e)}[/bold red]")
                                    print("[bold yellow]Restarting entire creation process...[/bold yellow]")
                                    continue  # Restart entire process
                                    
                        if r.status_code == 200:
                            jsn = r.json()
                            if 'signinName' in jsn:
                                email_address = jsn['signinName']
                                print(f"[green][+] Account created successfully: {email_address}[/green]")
                                cookies_str = ";".join([f"{key}={value}" for key, value in session.cookies.items()])
                                with open('out/accounts.txt', 'a', encoding='utf-8') as f:
                                    f.write(f"{email_address}|{password}|{cookies_str}\n")
                                
                                # Execute post-account creation processing
                                self.post_account_creation(email_address, password, proxy)
                            else:
                                print(f"[bold red]Account creation failed, response: {jsn}[/bold red]")
                                print("[bold yellow]Restarting entire creation process...[/bold yellow]")
                                continue  # Restart entire process
                        elif 'signinName' in response_json:
                            # Account created successfully, no captcha needed
                            email_address = response_json['signinName']
                            print(f"[green][+] Account created successfully: {email_address}[/green]")
                            cookies_str = ";".join([f"{key}={value}" for key, value in session.cookies.items()])
                            with open('out/accounts.txt', 'a', encoding='utf-8') as f:
                                f.write(f"{email_address}|{password}|{cookies_str}\n")
                            
                            # Execute post-account creation processing
                            self.post_account_creation(email_address, password, proxy)
                        else:
                            print(f"[bold red]Unclear response content: {response_json}[/bold red]")
                            print("[bold yellow]Restarting entire creation process...[/bold yellow]")
                            continue  # Restart entire process
                    else:
                        print(f"[bold red]Account creation failed, status code: {r.status_code}, response: {r.text}[/bold red]")
                        print("[bold yellow]Restarting entire creation process...[/bold yellow]")
                        continue  # Restart entire process
                        
                except Exception as e:
                    print(f"[bold red]Error in account creation request: {str(e)}[/bold red]")
                    print("[bold yellow]Restarting entire creation process...[/bold yellow]")
                    continue  # Restart entire process
                
                # Successfully completed entire process, break loop
                break
                
            except Exception as e:
                print(f"[bold red]Error occurred: {str(e)}[/bold red]")
                print("[yellow]Error details:[/yellow]")
                traceback.print_exc()
                print("[bold yellow]Restarting entire creation process...[/bold yellow]")
                continue  # Restart entire process
                
        # Reaching here means success or self.should_gen set to False
        return

    def post_account_creation(self, email, password, proxy=None):
        """
        Post-account creation processing: Enable IMAP and get OAuth token
        
        Args:
            email: Created email address
            password: Account password
            proxy: Proxy address (optional)
        """
        print(f"[bold magenta]Starting for {email} A) Enable IMAP service B) Get OAuth token[/bold magenta]")
        if proxy:
            print(f"[bold cyan]Using proxy: {proxy}[/bold cyan]")
        else:
            print("[bold yellow]No proxy provided, using direct connection[/bold yellow]")
        
        # Track IMAP and OAuth success status
        imap_success = False
        oauth_success = False
        refresh_token = None
        
        # IMAP and OAuth retry count settings
        max_retries = 10
        # Maximum proxy attempt count
        max_proxy_attempts = 3
        
        # A) Enable IMAP service (max 10 attempts)
        imap_retry_count = 0
        proxy_failure_count = 0  # Track proxy failure count
        
        while not imap_success and imap_retry_count < max_retries:
            try:
                if imap_retry_count > 0:
                    print(f"[bold blue]A) Attempt {imap_retry_count+1} to enable IMAP service...[/bold blue]")
                else:
                    print("[bold blue]A) Starting IMAP service enablement...[/bold blue]")
                
                # Check proxy availability
                if proxy:
                    proxy_available = check_proxy_available(proxy)
                    if not proxy_available:
                        proxy_failure_count += 1
                        if proxy_failure_count >= max_proxy_attempts:
                            print(f"[bold red]Tried {proxy_failure_count} proxies, all unavailable, switching to direct connection[/bold red]")
                            proxy = None
                        else:
                            print(f"[bold yellow]Current proxy unavailable, trying to get new proxy... (Attempt {proxy_failure_count}/{max_proxy_attempts})[/bold yellow]")
                            new_proxy = get_proxy_from_file()
                            if new_proxy:
                                proxy = new_proxy
                                proxy_available = check_proxy_available(proxy)
                                if proxy_available:
                                    print(f"[bold green]Successfully switched to new proxy: {proxy}[/bold green]")
                                    proxy_failure_count = 0  # Reset failure count
                                else:
                                    print(f"[bold yellow]New proxy also unavailable ({proxy_failure_count}/{max_proxy_attempts})[/bold yellow]")
                                    if proxy_failure_count >= max_proxy_attempts:
                                        print(f"[bold red]Tried {proxy_failure_count} proxies, all unavailable, switching to direct connection[/bold red]")
                                        proxy = None
                            else:
                                print(f"[bold red]Unable to get new proxy, will try direct connection[/bold red]")
                                proxy = None
                
                # Pass proxy parameter
                imap_result = enable_imap(email, password, proxy)
                if imap_result:
                    print(f"[green][+] IMAP service enabled successfully![/green]")
                    # Record IMAP enabled in account info
                    with open('out/accounts_with_imap.txt', 'a', encoding='utf-8') as f:
                        f.write(f"{email}|{password}|IMAP enabled\n")
                    imap_success = True
                else:
                    imap_retry_count += 1
                    print(f"[bold yellow][-] IMAP service enablement failed, tried {imap_retry_count}/{max_retries} times[/bold yellow]")
                    if imap_retry_count < max_retries:
                        print(f"[yellow]Waiting 3 seconds before retry...[/yellow]")
                        time.sleep(3)  # Wait before retry after failure
            except Exception as e:
                imap_retry_count += 1
                print(f"[bold red]Error enabling IMAP service: {str(e)}[/bold red]")
                traceback.print_exc()
                if imap_retry_count < max_retries:
                    print(f"[yellow]Waiting 3 seconds before retry...[/yellow]")
                    time.sleep(3)  # Wait before retry after error
        
        if not imap_success:
            print(f"[bold red]IMAP service enablement ultimately failed, tried {max_retries} times[/bold red]")
        
        # B) Get OAuth token (max 10 attempts)
        oauth_retry_count = 0
        proxy_failure_count = 0  # Reset proxy failure count
        
        if BLUEMAIL_AVAILABLE:
            # Use flatter loop structure instead of nested try-except
            while not oauth_success and oauth_retry_count < max_retries:
                oauth_retry_count += 1
                
                if oauth_retry_count > 1:
                    print(f"[bold blue]B) Attempt {oauth_retry_count} to get OAuth token...[/bold blue]")
                    time.sleep(3)  # Wait before retry
                else:
                    print("[bold blue]B) Starting OAuth token retrieval...[/bold blue]")
                
                try:
                    # Import module
                    from OauthBlueMail import BlueMailOAuthClient
                    print("[bold green]Successfully imported OAuth module[/bold green]")
                    
                    # Check proxy availability
                    if proxy:
                        proxy_available = check_proxy_available(proxy)
                        if not proxy_available:
                            proxy_failure_count += 1
                            if proxy_failure_count >= max_proxy_attempts:
                                print(f"[bold red]Tried {proxy_failure_count} proxies, all unavailable, switching to direct connection[/bold red]")
                                proxy = None
                            else:
                                print(f"[bold yellow]Current proxy unavailable, trying to get new proxy... (Attempt {proxy_failure_count}/{max_proxy_attempts})[/bold yellow]")
                                new_proxy = get_proxy_from_file()
                                if new_proxy:
                                    proxy = new_proxy
                                    proxy_available = check_proxy_available(proxy)
                                    if proxy_available:
                                        print(f"[bold green]Successfully switched to new proxy: {proxy}[/bold green]")
                                        proxy_failure_count = 0  # Reset failure count
                                    else:
                                        print(f"[bold yellow]New proxy also unavailable ({proxy_failure_count}/{max_proxy_attempts})[/bold yellow]")
                                        if proxy_failure_count >= max_proxy_attempts:
                                            print(f"[bold red]Tried {proxy_failure_count} proxies, all unavailable, switching to direct connection[/bold red]")
                                            proxy = None
                                else:
                                    print(f"[bold red]Unable to get new proxy, will try direct connection[/bold red]")
                                    proxy = None
                    
                    # Create client and set proxy
                    oauth_client = BlueMailOAuthClient(email, password)
                    if proxy:
                        proxy_set = oauth_client.set_proxy(proxy)
                        if proxy_set:
                            print(f"[cyan]OAuth client proxy set: {proxy}[/cyan]")
                        else:
                            print(f"[yellow]OAuth client proxy setting failed, will use direct connection[/yellow]")
                    else:
                        print("[yellow]No proxy provided, OAuth will use direct connection[/yellow]")
                    
                    # Execute OAuth flow
                    refresh_token = oauth_client.execute_auth_flow()
                    if refresh_token:
                        print(f"[green][+] OAuth token retrieval successful![/green]")
                        
                        # Don't save OAuth data to file - simulating original save_data=False effect
                        # If data saving is needed, uncomment the following line
                        # oauth_client.save_auth_data()
                        
                        # Record account with successful OAuth token
                        with open('out/accounts_with_oauth.txt', 'a', encoding='utf-8') as f:
                            f.write(f"{email}|{password}|OAuth successful\n")
                        oauth_success = True
                        break  # Break loop on success
                        
                except ImportError as e:
                    print(f"[bold red]Failed to import OauthBlueMail module: {str(e)}[/bold red]")
                    traceback.print_exc()
                    break  # Module import failed, stop trying
                except Exception as e:
                    print(f"[bold red]Error getting OAuth token: {str(e)}[/bold red]")
                    traceback.print_exc()
                    # Continue loop retry
            
            if not oauth_success:
                print(f"[bold red]OAuth token retrieval ultimately failed, tried {oauth_retry_count}/{max_retries} times[/bold red]")
        else:
            print("[bold yellow][-] BlueMail OAuth module not loaded, skipping OAuth token retrieval[/bold yellow]")
            
        # C) If both IMAP and OAuth successful, upload account info to API
        if imap_success and oauth_success and refresh_token:
            print(f"[bold blue]IMAP and OAuth successful...[/bold blue]")
            with open("out/accounts_with_refresh_token.txt", "a", encoding="utf-8") as f:
                f.write(f"{email}:{password}:{refresh_token}:8b4ba9dd-3ea5-4e5f-86f1-ddba2230dcf2\n")
        elif imap_success and oauth_success and not refresh_token:
            print("[bold yellow][-] Refresh token empty, skipping API upload[/bold yellow]")
        elif not (imap_success and oauth_success):
            print("[bold yellow][-] IMAP or OAuth not successful, skipping API upload[/bold yellow]")

def start_multiprocess(thread_count: int, domain_index: int = None):
    """
    Start multi-threaded continuous account creation
    
    Args:
        thread_count: Number of threads to run
        domain_index: Domain index to use, None means random
    """
    print(f"[bold blue]Starting continuous account creation with {thread_count} threads[/bold blue]")
    
    def continuous_account_creation():
        outlook = Outlook()
        while True:
            try:
                outlook.init_session(domain_index)
            except Exception as e:
                print(f"[bold red]Error in thread: {str(e)}[/bold red]")
                print("[yellow]Restarting account creation process...[/yellow]")
                time.sleep(3)  # Wait before retry after error
                continue

    # Start specified number of threads
    for i in range(thread_count):
        print(f"[blue]Starting thread {i+1}/{thread_count}[/blue]")
        threading.Thread(target=continuous_account_creation, daemon=True).start()
        
    # Keep main thread alive
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n[bold yellow]Stopping all threads...[/bold yellow]")
        sys.exit(0)
            

if __name__ == '__main__':
    # First print welcome message
    logger.info("=" * 50)
    logger.info("===== Microsoft Outlook Automatic Account Creation Tool =====")
    logger.info("=" * 50)
    
    # Add important notes
    logger.info("Note: Please ensure KoCaptcha API token and valid proxy API are properly configured")
    
    # Test proxy API connection
    logger.info("Testing proxy API connection...")
    proxy = get_proxy_from_file()
    if proxy:
        logger.info(f"Proxy API connection successful! Got proxy: {proxy}")
    else:
        logger.warning("Unable to get proxy, please check proxy API configuration. Will continue but may affect success rate.")
    
    # Select domain
    print("\nSupported Outlook domains:")
    for i, (name, tld, mkt) in enumerate(OUTLOOK_DOMAINS, 1):
        domain = f"{name}.{tld}"
        print(f"{i:2d}. {domain:<20} (Region: {mkt})")
        
    print("\nPlease enter a number to select domain, or enter 0 for random selection")
    domain_choice = None
    
    try:
        choice = int(input("Please select domain number to use (1-32, or enter 0 for random): "))
        if 1 <= choice <= len(OUTLOOK_DOMAINS):
            domain_choice = choice - 1  # Convert to 0-based index
            domain_name, domain_tld, _ = OUTLOOK_DOMAINS[domain_choice]
            logger.info(f"Selected domain: {domain_name}.{domain_tld}")
        elif choice == 0:
            domain_choice = None
            logger.info("Will randomly select domain")
        else:
            logger.warning(f"Invalid choice: {choice}, will randomly select domain")
    except ValueError:
        logger.warning("Invalid input, will randomly select domain")
    
    # Set thread count
    try:
        thread_count = int(input("Please enter number of threads to run: "))
        if thread_count <= 0:
            thread_count = 1
            logger.warning("Thread count must be greater than 0, set to 1")
    except ValueError:
        thread_count = 1
        logger.warning("Invalid input, thread count set to 1")
    
    # Start continuous account creation
    start_multiprocess(thread_count, domain_choice)
