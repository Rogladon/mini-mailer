import {
  Modal,
  ModalOverlay,
  ModalContent,
  ModalHeader,
  ModalCloseButton,
  ModalBody,
  ModalFooter,
  Button,
} from '@chakra-ui/react';

interface Props {
  isOpen: boolean
  onClose: () => void
  previewContent: string
}

const Preview = ({ isOpen, onClose, previewContent }: Props) => {
  return <Modal isOpen={isOpen} onClose={onClose} size="xl">
    <ModalOverlay />
    <ModalContent>
      <ModalHeader>Предпросмотр письма</ModalHeader>
      <ModalCloseButton />
      <ModalBody>
        <div
          style={{ padding: '10px', border: '1px solid #ccc', minHeight: '200px' }}
          dangerouslySetInnerHTML={{ __html: previewContent }}
        />
      </ModalBody>
      <ModalFooter>
        <Button colorScheme="teal" onClick={onClose}>
          Закрыть
        </Button>
      </ModalFooter>
    </ModalContent>
  </Modal>
}

export default Preview
