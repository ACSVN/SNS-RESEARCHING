$strPath = (Convert-Path ..)

#$path = "C:\SNS-TWITTER\images\"                     # �o�͐�
$Path = $strPath + "\images\"
$file = "Cap.png"             # {0:000} 0�Ԗڂ̈�����3����0���ߐ����Ƃ���B
Add-Type -A System.Windows.Forms      # ���̃Z�b�V�����Ɍ^��ǉ��B
[Windows.Forms.Screen]::PrimaryScreen.Bounds |                   # �f�B�X�v���C�T�C�Y���擾���A
    % {[Drawing.Bitmap]::New($_.Width, $_.Height)} |             # ���̃T�C�Y��Bitmap�f�[�^�����B
    sv b                                                         # Set-Variable �ϐ��ւ̑���Bb��Bitmap�I�u�W�F�N�g�����B

[Drawing.Graphics]::FromImage($b).CopyFromScreen(0, 0, $b.Size)  # �`��I�u�W�F�N�g�ŉ�ʃL���v�`�����āABitmap�f�[�^�֔��f�B

md -f $path                           # �t�H���_�쐬�B-f:force ���ɑ��݂���ꍇ�G���[��}�~�B

1..1000 |                             # 1�`1000�܂�
    % { $path + $file -f $_ }  |      # �J��Ԃ��A�t�H�[�}�b�g�Bc:\tmp\Cap_001.png�Ȃǂɕϊ��B
    ? { (Test-Path $_) -eq $false } | # ���̓��A�t�@�C�������݂��Ȃ��p�X�݂̂��擾�B
    select -F 1 |                     # �ŏ���1���݂̂Ƀt�B���^���āA
    % { $b.Save($_) }                 # ���̃p�X��Bitmap�f�[�^���t�@�C���ۑ��B