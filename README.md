<!-- Improved compatibility of back to top link: See: https://github.com/othneildrew/Best-README-Template/pull/73 -->
<a name="readme-top"></a>
<!--
*** Thanks for checking out the Best-README-Template. If you have a suggestion
*** that would make this better, please fork the repo and create a pull request
*** or simply open an issue with the tag "enhancement".
*** Don't forget to give the project a star!
*** Thanks again! Now go create something AMAZING! :D
-->



<!-- PROJECT SHIELDS -->
<!--
*** I'm using markdown "reference style" links for readability.
*** Reference links are enclosed in brackets [ ] instead of parentheses ( ).
*** See the bottom of this document for the declaration of the reference variables
*** for contributors-url, forks-url, etc. This is an optional, concise syntax you may use.
*** https://www.markdownguide.org/basic-syntax/#reference-style-links
-->
[![Contributors][contributors-shield]][contributors-url]
[![Issues][issues-shield]][issues-url]
[![LinkedIn][linkedin-shield]][linkedin-url]



<!-- PROJECT LOGO -->
<br />
<div align="center">
  <a href="https://github.com/cordwainersmith/LocalCertMan">
    <img src="images/logo.png" alt="Logo" width="80" height="80">
  </a>

<h3 align="center">LocalCertMan</h3>

  <p align="center">
a compiled PowerShell script that uses the Windows Forms framework to create a graphical user interface that performs different certificate management tasks on Windows,in particular:
certificates lists export, certificate installation, certificate conversion, and certificate binding to an IIS website.
    <br />
    <br />
    <a href="https://github.com/cordwainersmith/LocalCertMan/issues">Report Bug</a>
    ·
    <a href="mailto:cordwainer@AlphaRalpha.net">Request Feature</a>
  </p>
</div>

[![Product Name Screen Shot][product-screenshot]](https://example.com)



<!-- ABOUT THE PROJECT -->
## Features

* Displaying the details of all the installed Personal and Root certificates, detects certificates that are about to expire and highlights them.
* Exporting certificates lists (Personal and Trusted Root CA) along with the DNS names and port binding to CSV files 
* Installing a password protected PFX Personal certificate
* Installing a trusted root certificate to the local machine certificate store
* Merging CER and KEY files to a passsword protected PFX
* Bind a certificate to the IIS default website
* View certificate details from URL
<p align="right">(<a href="#readme-top">back to top</a>)</p>


<!-- Prerequisites -->

## Prerequisites

PowerShell: 
* requires PowerShell version 5.1 or later to be installed on the system
* requires PowerShell WebAdministration module to be installed to allow SSL bindings
* Execution policy: By default, the PowerShell execution policy is set to Restricted, which prevents running scripts. You will need to set the execution policy to RemoteSigned or Unrestricted to run this tool.

.NET Framework:
* requires version 4.5 or later of the .NET Framework to be installed on the system.

Run As Administrator Permisssions:
* the tool needs access to the LocalMachine certificates store to be able to export certificates and add the root CA to the trust store, and should be run with administrative privileges.


<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- LICENSE -->
## License

Distributed under the MIT License. See `LICENSE.txt` for more information.

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- CONTACT -->
## Contact

cordwainer@alpharalpha.net

See the [open issues](https://github.com/cordwainersmith/LocalCertMan/issues) for a full list of proposed features (and known issues).

Project Link: [https://github.com/cordwainersmith/LocalCertMan](https://github.com/cordwainersmith/LocalCertMan)

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- ACKNOWLEDGMENTS -->


<!-- MARKDOWN LINKS & IMAGES -->
<!-- https://www.markdownguide.org/basic-syntax/#reference-style-links -->
[contributors-shield]: https://img.shields.io/github/contributors/cordwainersmith/LocalCertMan.svg?style=for-the-badge
[contributors-url]: https://github.com/cordwainersmith/LocalCertMan/graphs/contributors
[forks-shield]: https://img.shields.io/github/forks/cordwainersmith/LocalCertMan.svg?style=for-the-badge
[forks-url]: https://github.com/cordwainersmith/LocalCertMan/network/members
[stars-shield]: https://img.shields.io/github/stars/cordwainersmith/LocalCertMan.svg?style=for-the-badge
[stars-url]: https://github.com/cordwainersmith/LocalCertMan/stargazers
[issues-shield]: https://img.shields.io/github/issues/cordwainersmith/LocalCertMan.svg?style=for-the-badge
[issues-url]: https://github.com/cordwainersmith/LocalCertMan/issues
[license-shield]: https://img.shields.io/github/license/cordwainersmith/LocalCertMan.svg?style=for-the-badge
[license-url]: https://github.com/cordwainersmith/LocalCertMan/blob/master/LICENSE.txt
[linkedin-shield]: https://img.shields.io/badge/-LinkedIn-black.svg?style=for-the-badge&logo=linkedin&colorB=555
[linkedin-url]: https://linkedin.com/in/liran-baba-77b4b421
[product-screenshot]: images/screenshot.png
[Next.js]: https://img.shields.io/badge/next.js-000000?style=for-the-badge&logo=nextdotjs&logoColor=white
[Next-url]: https://nextjs.org/
[React.js]: https://img.shields.io/badge/React-20232A?style=for-the-badge&logo=react&logoColor=61DAFB
[React-url]: https://reactjs.org/
[Vue.js]: https://img.shields.io/badge/Vue.js-35495E?style=for-the-badge&logo=vuedotjs&logoColor=4FC08D
[Vue-url]: https://vuejs.org/
[Angular.io]: https://img.shields.io/badge/Angular-DD0031?style=for-the-badge&logo=angular&logoColor=white
[Angular-url]: https://angular.io/
[Svelte.dev]: https://img.shields.io/badge/Svelte-4A4A55?style=for-the-badge&logo=svelte&logoColor=FF3E00
[Svelte-url]: https://svelte.dev/
[Laravel.com]: https://img.shields.io/badge/Laravel-FF2D20?style=for-the-badge&logo=laravel&logoColor=white
[Laravel-url]: https://laravel.com
[Bootstrap.com]: https://img.shields.io/badge/Bootstrap-563D7C?style=for-the-badge&logo=bootstrap&logoColor=white
[Bootstrap-url]: https://getbootstrap.com
[JQuery.com]: https://img.shields.io/badge/jQuery-0769AD?style=for-the-badge&logo=jquery&logoColor=white
[JQuery-url]: https://jquery.com 
