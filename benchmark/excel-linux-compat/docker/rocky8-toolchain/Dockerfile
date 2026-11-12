FROM rockylinux@sha256:9794037624aaa6212aeada1d28861ef5e0a935adaf93e4ef79837119f2a2d04c

RUN cat > /etc/yum.repos.d/rocky8-direct.repo <<'EOF'
[rocky8-baseos]
name=Rocky Linux 8 BaseOS (direct mirror)
baseurl=https://dl.rockylinux.org/pub/rocky/8/BaseOS/x86_64/os/
enabled=1
gpgcheck=1
gpgkey=file:///etc/pki/rpm-gpg/RPM-GPG-KEY-rockyofficial

[rocky8-appstream]
name=Rocky Linux 8 AppStream (direct mirror)
baseurl=https://dl.rockylinux.org/pub/rocky/8/AppStream/x86_64/os/
enabled=1
gpgcheck=1
gpgkey=file:///etc/pki/rpm-gpg/RPM-GPG-KEY-rockyofficial
EOF
RUN dnf --disablerepo='*' --enablerepo=rocky8-baseos --enablerepo=rocky8-appstream install -y \
    gcc-c++ cmake make binutils file fontconfig \
    && dnf clean all

RUN dnf --disablerepo='*' --enablerepo=rocky8-baseos --enablerepo=rocky8-appstream install -y libicu \
    && dnf clean all

RUN dnf --disablerepo='*' --enablerepo=rocky8-baseos --enablerepo=rocky8-appstream install -y rust cargo \
    && dnf clean all
